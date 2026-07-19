'use strict';

/**
 * TBH 通关时间探针 — 托盘程序用
 *
 * 通关 toast（无难度文字）:
 *   通关了<color=...>关卡 2-3</color>。(48秒) <...>[11:53]</...>
 *
 * 难度来源（游戏内 0~3）:
 *   UI_Portal.m_currentStageDifficulty
 *   0=普通 1=噩梦 2=地狱 3=折磨
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
var il2cpp_class_get_field_from_name = api('il2cpp_class_get_field_from_name', 'pointer', ['pointer', 'pointer']);
var il2cpp_field_get_offset = api('il2cpp_field_get_offset', 'int', ['pointer']);
var il2cpp_class_get_methods = api('il2cpp_class_get_methods', 'pointer', ['pointer', 'pointer']);
var il2cpp_method_get_name = api('il2cpp_method_get_name', 'pointer', ['pointer']);

function cstr(s) {
  return Memory.allocUtf8String(s);
}

function readCString(p) {
  try {
    if (!p || p.isNull()) return '';
    return p.readUtf8String() || '';
  } catch (e) {
    return '';
  }
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

function isReadable(p) {
  try {
    if (!p || p.isNull()) return false;
    var r = Process.findRangeByAddress(p);
    return !!(r && r.protection.indexOf('r') !== -1);
  } catch (e) {
    return false;
  }
}

function nowStamp() {
  var d = new Date();
  function p(n) {
    return n < 10 ? '0' + n : '' + n;
  }
  return p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}

function emit(kind, payload) {
  payload = payload || {};
  payload.kind = kind;
  payload.ts = nowStamp();
  send(payload);
}

// ---------- 难度 ----------
// 0=普通 1=噩梦 2=地狱 3=折磨（与 TBH-DropMonitor / UI_Portal 一致）
var DIFF_NAMES = ['普通', '噩梦', '地狱', '折磨'];
var g_portal = ptr(0);
var g_diffOffset = -1;
var g_lastDiffId = -1;
var g_lastDiffName = '未知';
var g_portalHooks = 0;

function difficultyName(id) {
  id = parseInt(id, 10);
  if (id >= 0 && id <= 3) return DIFF_NAMES[id];
  return '未知';
}

function parseDifficultyFromText(text) {
  var t = stripRichText(text || '');
  if (!t) return null;
  // 纯按钮/页签
  if (t === '普通' || t === '噩梦' || t === '地狱' || t === '折磨') return t;
  // 带前缀：折磨 2-9 / [折磨]关卡
  if (t.indexOf('折磨') >= 0) return '折磨';
  if (t.indexOf('地狱') >= 0) return '地狱';
  if (t.indexOf('噩梦') >= 0) return '噩梦';
  // 「普通」太泛，仅整词匹配
  if (/(^|[^a-zA-Z\u4e00-\u9fff])普通([^a-zA-Z\u4e00-\u9fff]|$)/.test(t) && t.length <= 12) return '普通';
  return null;
}

function difficultyIdFromName(name) {
  for (var i = 0; i < DIFF_NAMES.length; i++) {
    if (DIFF_NAMES[i] === name) return i;
  }
  return -1;
}

function rememberDifficulty(id, name, source) {
  if (typeof id === 'number' && id >= 0 && id <= 3) {
    g_lastDiffId = id;
    g_lastDiffName = difficultyName(id);
    return;
  }
  if (name && DIFF_NAMES.indexOf(name) >= 0) {
    g_lastDiffName = name;
    g_lastDiffId = difficultyIdFromName(name);
  }
}

function readPortalDifficulty() {
  if (g_diffOffset < 0) return null;
  if (!isReadable(g_portal)) return null;
  try {
    var v = g_portal.add(g_diffOffset).readS32();
    if (v >= 0 && v <= 3) {
      rememberDifficulty(v, null, 'portal');
      return { id: v, name: difficultyName(v), source: 'UI_Portal' };
    }
  } catch (e) {}
  return null;
}

function currentDifficulty() {
  var fromPortal = readPortalDifficulty();
  if (fromPortal) return fromPortal;
  if (g_lastDiffId >= 0) {
    return { id: g_lastDiffId, name: g_lastDiffName || difficultyName(g_lastDiffId), source: 'cached' };
  }
  if (g_lastDiffName && g_lastDiffName !== '未知') {
    return { id: difficultyIdFromName(g_lastDiffName), name: g_lastDiffName, source: 'cached-name' };
  }
  return { id: -1, name: '未知', source: 'none' };
}

function setupPortalDifficultyCapture() {
  var klass = findClass('TaskbarHero.UI', 'UI_Portal');
  if (!klass || klass.isNull()) {
    emit('status', { text: 'UI_Portal class not found' });
    return;
  }
  try {
    var field = il2cpp_class_get_field_from_name(klass, cstr('m_currentStageDifficulty'));
    if (field && !field.isNull()) {
      g_diffOffset = il2cpp_field_get_offset(field);
      emit('status', { text: 'UI_Portal.m_currentStageDifficulty offset=0x' + g_diffOffset.toString(16) });
    } else {
      emit('status', { text: 'm_currentStageDifficulty field not found' });
    }
  } catch (e) {
    emit('status', { text: 'field resolve failed: ' + e });
  }

  // 捕获 portal 实例：hook 若干实例方法的 this
  try {
    var iter = Memory.alloc(Process.pointerSize);
    iter.writePointer(ptr(0));
    var n = 0;
    while (n < 120 && g_portalHooks < 40) {
      var method = il2cpp_class_get_methods(klass, iter);
      if (!method || method.isNull()) break;
      var name = readCString(il2cpp_method_get_name(method));
      n++;
      if (!name || name === '.ctor' || name === '.cctor') continue;
      var fp = methodPtr(method);
      if (!fp || fp.isNull()) continue;
      try {
        Interceptor.attach(fp, {
          onEnter: function (args) {
            try {
              if (args[0] && !args[0].isNull() && isReadable(args[0])) {
                g_portal = args[0];
                readPortalDifficulty();
              }
            } catch (e0) {}
          }
        });
        g_portalHooks++;
      } catch (e1) {}
    }
  } catch (e) {
    emit('status', { text: 'portal instance hook failed: ' + e });
  }
  emit('status', { text: 'UI_Portal instance hooks=' + g_portalHooks });
}

// ---------- 通关解析 ----------
var RE_CLEAR = /通关[了]?\s*关卡\s*(\d+)\s*[-－—~～]\s*(\d+)\s*[。\.]?\s*[\(（]\s*(\d+)\s*秒\s*[\)）](?:\s*[\[【](\d{1,2}:\d{2})[\]】])?/;

var g_seenKeys = {};
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

  // 从 UI 文案侧路更新难度（页签等）
  var diffName = parseDifficultyFromText(text);
  if (diffName && text.indexOf('通关') < 0) {
    rememberDifficulty(-1, diffName, source);
  }

  if (text.indexOf('通关') < 0) return;

  var parsed = parseClearNotice(text);
  if (!parsed) return;

  var diff = currentDifficulty();
  // toast 自身若带难度字样，优先用
  var inText = parseDifficultyFromText(parsed.raw);
  if (inText) {
    diff = { id: difficultyIdFromName(inText), name: inText, source: 'toast-text' };
    rememberDifficulty(diff.id, diff.name, 'toast');
  }

  // 唯一键：难度 + 关卡 + 秒数 + 通知时钟
  var key =
    (diff.name || '未知') +
    '|' +
    parsed.stage +
    '|' +
    parsed.clearSeconds +
    '|' +
    (parsed.noticeTime || '');
  if (g_seenKeys[key]) return;
  g_seenKeys[key] = 1;
  var keys = Object.keys(g_seenKeys);
  if (keys.length > 500) {
    for (var i = 0; i < keys.length - 400; i++) delete g_seenKeys[keys[i]];
  }

  emit('clear_time', {
    source: source,
    stage: parsed.stage,
    chapter: parsed.chapter,
    level: parsed.level,
    clearSeconds: parsed.clearSeconds,
    noticeTime: parsed.noticeTime,
    difficulty: diff.name || '未知',
    difficultyId: typeof diff.id === 'number' ? diff.id : -1,
    difficultySource: diff.source || 'none',
    raw: parsed.raw,
    dedupeKey: key
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
  setupPortalDifficultyCapture();

  var targets = [
    { ns: 'TMPro', klass: 'TMP_Text', name: 'SetText', argcs: [1, 2, 3, 4] },
    { ns: 'TMPro', klass: 'TextMeshProUGUI', name: 'SetText', argcs: [1, 2, 3, 4] },
    { ns: 'TMPro', klass: 'TMP_Text', name: 'set_text', argcs: [1] },
    { ns: 'UnityEngine.UI', klass: 'Text', name: 'set_text', argcs: [1] }
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
  emit('status', {
    text: 'hooks=' + g_hooked + ' portalHooks=' + g_portalHooks + ' diffOffset=' + g_diffOffset,
    hooked: g_hooked
  });
}

rpc.exports = {
  stats: function () {
    var d = currentDifficulty();
    return {
      hooked: g_hooked,
      portalHooks: g_portalHooks,
      diffOffset: g_diffOffset,
      difficulty: d.name,
      difficultyId: d.id
    };
  },
  getDifficulty: function () {
    return currentDifficulty();
  }
};

setImmediate(function () {
  try {
    installHooks();
  } catch (e) {
    emit('fatal', { error: '' + e });
  }
});
