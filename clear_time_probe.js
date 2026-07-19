'use strict';

/**
 * TBH 通关时间探针 — 托盘程序用
 * 实际 TMP: 通关了<color=...>关卡 2-3</color>。(48秒) <...>[11:53]</...>
 */

var GA = Process.enumerateModules().find(function (m) {
  return m.name.toLowerCase().indexOf('gameassembly') !== -1;
});
if (!GA) throw new Error('GameAssembly.dll not found');

function api(name, ret, args) {
  return new NativeFunction(GA.getExportByName(name), ret, args);
}

var il2cpp_domain_get = api('il2cpp_domain_get', 'pointer', []);
var il2cpp_domain_get_assemblies = api('il2cpp_domain_get_assemblies', 'pointer', ['pointer', 'pointer']);
var il2cpp_assembly_get_image = api('il2cpp_assembly_get_image', 'pointer', ['pointer']);
var il2cpp_class_from_name = api('il2cpp_class_from_name', 'pointer', ['pointer', 'pointer', 'pointer']);
var il2cpp_class_get_method_from_name = api('il2cpp_class_get_method_from_name', 'pointer', ['pointer', 'pointer', 'int']);

function cstr(s) {
  return Memory.allocUtf8String(s);
}

function methodPtr(methodInfo) {
  try {
    if (!methodInfo || methodInfo.isNull()) return ptr(0);
    var fp = methodInfo.readPointer();
    if (!fp || fp.isNull()) return ptr(0);
    return fp;
  } catch (e) {
    return ptr(0);
  }
}

function findClass(ns, name) {
  var domain = il2cpp_domain_get();
  var countPtr = Memory.alloc(4);
  var assemblies = il2cpp_domain_get_assemblies(domain, countPtr);
  var count = countPtr.readU32();
  for (var i = 0; i < count; i++) {
    var asm = assemblies.add(i * Process.pointerSize).readPointer();
    if (!asm || asm.isNull()) continue;
    var image = il2cpp_assembly_get_image(asm);
    if (!image || image.isNull()) continue;
    var klass = il2cpp_class_from_name(image, cstr(ns), cstr(name));
    if (klass && !klass.isNull()) return klass;
  }
  return ptr(0);
}

function findMethod(ns, klassName, methodName, argc) {
  var klass = findClass(ns, klassName);
  if (!klass || klass.isNull()) return ptr(0);
  return methodPtr(il2cpp_class_get_method_from_name(klass, cstr(methodName), argc));
}

function readIl2CppString(strObj) {
  try {
    if (!strObj || strObj.isNull()) return '';
    var len = strObj.add(0x10).readS32();
    if (len <= 0 || len > 4096) return '';
    return strObj.add(0x14).readUtf16String(len) || '';
  } catch (e) {
    return '';
  }
}

function nowStamp() {
  var d = new Date();
  function p(n) {
    return n < 10 ? '0' + n : '' + n;
  }
  return (
    p(d.getHours()) +
    ':' +
    p(d.getMinutes()) +
    ':' +
    p(d.getSeconds())
  );
}

function emit(kind, payload) {
  payload = payload || {};
  payload.kind = kind;
  payload.ts = nowStamp();
  send(payload);
}

// 去标签后: 通关了关卡 2-3。(48秒) [11:53]
var RE_CLEAR = /通关[了]?\s*关卡\s*(\d+)\s*[-－—~～]\s*(\d+)\s*[。\.]?\s*[\(（]\s*(\d+)\s*秒\s*[\)）](?:\s*[\[【](\d{1,2}:\d{2})[\]】])?/;

var g_recent = {};
var RECENT_MS = 2500;
var g_hooked = 0;

function stripRichText(text) {
  return String(text || '')
    .replace(/<[^>]+>/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseClearNotice(text) {
  if (!text || text.indexOf('通关') < 0) return null;
  var plain = stripRichText(text);
  plain = plain.replace(/（/g, '(').replace(/）/g, ')').replace(/【/g, '[').replace(/】/g, ']');
  var m = plain.match(RE_CLEAR);
  if (!m) return null;
  return {
    stage: m[1] + '-' + m[2],
    chapter: parseInt(m[1], 10),
    level: parseInt(m[2], 10),
    clearSeconds: parseInt(m[3], 10),
    noticeTime: m[4] || '',
    raw: plain
  };
}

function handleText(source, text) {
  if (!text || typeof text !== 'string') return;
  if (text.indexOf('通关') < 0) return;

  var parsed = parseClearNotice(text);
  if (!parsed) return;

  // 去重：同关卡+秒数+通知时钟 在短时间内只报一次
  var key = parsed.stage + '|' + parsed.clearSeconds + '|' + parsed.noticeTime;
  var now = Date.now();
  if (g_recent[key] && now - g_recent[key] < RECENT_MS) return;
  g_recent[key] = now;
  if (Object.keys(g_recent).length > 200) g_recent = {};

  emit('clear_time', {
    source: source,
    stage: parsed.stage,
    chapter: parsed.chapter,
    level: parsed.level,
    clearSeconds: parsed.clearSeconds,
    noticeTime: parsed.noticeTime,
    raw: parsed.raw
  });
}

function tryHook(fp, label) {
  if (!fp || fp.isNull()) return false;
  try {
    Interceptor.attach(fp, {
      onEnter: function (args) {
        try {
          for (var i = 1; i <= 3; i++) {
            var s = readIl2CppString(args[i]);
            if (s) handleText(label, s);
          }
        } catch (e) {}
      }
    });
    g_hooked++;
    return true;
  } catch (e) {
    return false;
  }
}

function installHooks() {
  var targets = [
    { ns: 'TMPro', klass: 'TMP_Text', name: 'SetText', argcs: [1, 2, 3, 4] },
    { ns: 'TMPro', klass: 'TextMeshProUGUI', name: 'SetText', argcs: [1, 2, 3, 4] },
    { ns: 'TMPro', klass: 'TMP_Text', name: 'set_text', argcs: [1] }
  ];
  var seen = {};
  targets.forEach(function (t) {
    t.argcs.forEach(function (argc) {
      var fp = findMethod(t.ns, t.klass, t.name, argc);
      if (!fp || fp.isNull()) return;
      var id = fp.toString() + '|' + t.name;
      if (seen[id]) return;
      seen[id] = true;
      tryHook(fp, t.klass + '.' + t.name + '/' + argc);
    });
  });
  emit('status', { text: 'hooks=' + g_hooked, hooked: g_hooked });
}

rpc.exports = {
  stats: function () {
    return { hooked: g_hooked };
  }
};

setImmediate(function () {
  try {
    installHooks();
  } catch (e) {
    emit('fatal', { error: '' + e });
  }
});
