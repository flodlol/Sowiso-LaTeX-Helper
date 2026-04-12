const latexInput = document.getElementById("latexInput");
const insertMode = document.getElementById("insertMode");
const sowisoConvert = document.getElementById("sowisoConvert");
const previewButton = document.getElementById("previewButton");
const previewImage = document.getElementById("previewImage");
const previewPlaceholder = document.getElementById("previewPlaceholder");
const insertButton = document.getElementById("insertButton");
const clearButton = document.getElementById("clearButton");
const statusEl = document.getElementById("status");
const copyDebugButton = document.getElementById("copyDebugButton");
const clearDebugButton = document.getElementById("clearDebugButton");
const debugLogEl = document.getElementById("debugLog");
const themeChips = [...document.querySelectorAll(".theme-chip")];
const systemTheme = window.matchMedia("(prefers-color-scheme: dark)");

const DEFAULT_SETTINGS = {
  themeMode: "auto"
};

let currentThemeMode = "auto";
let debugLines = [];

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", isError);
}

function safeStringify(value) {
  try {
    return JSON.stringify(value);
  } catch (_error) {
    return String(value);
  }
}

function appendDebug(message, payload) {
  const timestamp = new Date().toISOString().slice(11, 23);
  const line = payload === undefined ? `[${timestamp}] ${message}` : `[${timestamp}] ${message} ${safeStringify(payload)}`;
  debugLines.push(line);

  if (debugLines.length > 200) {
    debugLines = debugLines.slice(-200);
  }

  if (debugLogEl) {
    debugLogEl.textContent = debugLines.join("\n");
    debugLogEl.scrollTop = debugLogEl.scrollHeight;
  }

  console.log("[SowisoHelper]", message, payload ?? "");
}

async function copyDebugLog() {
  const text = debugLines.join("\n");
  if (!text) {
    setStatus("Debug log is empty.", true);
    return;
  }

  try {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
      await navigator.clipboard.writeText(text);
    } else {
      const helper = document.createElement("textarea");
      helper.value = text;
      helper.setAttribute("readonly", "readonly");
      helper.style.position = "fixed";
      helper.style.opacity = "0";
      document.body.append(helper);
      helper.select();
      document.execCommand("copy");
      helper.remove();
    }
    setStatus("Debug log copied.", false);
    appendDebug("Debug log copied");
  } catch (error) {
    setStatus("Could not copy debug log.", true);
    appendDebug("Failed to copy debug log", { error: error && error.message ? error.message : String(error) });
  }
}

function getResolvedTheme(mode) {
  if (mode === "light" || mode === "dark") {
    return mode;
  }

  return systemTheme.matches ? "dark" : "light";
}

function setActiveThemeChip(mode) {
  for (const chip of themeChips) {
    chip.classList.toggle("active", chip.dataset.theme === mode);
  }
}

function applyTheme(mode) {
  currentThemeMode = mode;
  document.documentElement.dataset.theme = getResolvedTheme(mode);
  setActiveThemeChip(mode);
}

function persistSettings(settingsPatch) {
  chrome.storage.local.get(DEFAULT_SETTINGS, (current) => {
    const next = { ...current, ...settingsPatch };
    chrome.storage.local.set(next);
  });
}

function wrapFormula(raw, mode) {
  if (mode === "inline") {
    return `$${raw}$`;
  }

  if (mode === "block") {
    return `$$${raw}$$`;
  }

  return raw;
}

function findMatchingBrace(input, openIndex) {
  if (input[openIndex] !== "{") {
    return -1;
  }

  let depth = 0;
  for (let i = openIndex; i < input.length; i += 1) {
    const ch = input[i];
    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) return i;
    }
  }

  return -1;
}

function replaceFracLatex(input) {
  let value = input;
  let guard = 0;

  while (value.includes("\\frac") && guard < 200) {
    guard += 1;
    const idx = value.indexOf("\\frac");
    const numOpen = value.indexOf("{", idx + 5);
    if (numOpen === -1) break;

    const numClose = findMatchingBrace(value, numOpen);
    if (numClose === -1) break;

    const denOpen = value.indexOf("{", numClose + 1);
    if (denOpen === -1) break;

    const denClose = findMatchingBrace(value, denOpen);
    if (denClose === -1) break;

    const num = value.slice(numOpen + 1, numClose);
    const den = value.slice(denOpen + 1, denClose);
    const replacement = `(${num})/(${den})`;
    value = `${value.slice(0, idx)}${replacement}${value.slice(denClose + 1)}`;
  }

  return value;
}

function replaceSqrtLatex(input) {
  let value = input;
  let guard = 0;

  while (value.includes("\\sqrt") && guard < 200) {
    guard += 1;
    const idx = value.indexOf("\\sqrt");
    const argOpen = value.indexOf("{", idx + 5);
    if (argOpen === -1) break;

    const argClose = findMatchingBrace(value, argOpen);
    if (argClose === -1) break;

    const arg = value.slice(argOpen + 1, argClose);
    const replacement = `sqrt(${arg})`;
    value = `${value.slice(0, idx)}${replacement}${value.slice(argClose + 1)}`;
  }

  return value;
}

function isAsciiLetter(ch) {
  return /^\p{L}$/u.test(ch);
}

function isIdentifierToken(token) {
  return /^[\p{L}]+$/u.test(token);
}

function isNumberToken(token) {
  return /^[0-9]+(?:\.[0-9]+)?$/.test(token);
}

function isOperatorToken(token) {
  return /^[+\-*/^=<>!&,|]$/.test(token) || token === "<=" || token === ">=" || token === "!=" || token === "_";
}

function isScriptToken(token) {
  return typeof token === "string" && (token.startsWith("_") || token.startsWith("^"));
}

function tokenizeLinearExpression(input) {
  const tokens = [];
  let i = 0;

  while (i < input.length) {
    const ch = input[i];

    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }

    const two = input.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "!=") {
      tokens.push(two);
      i += 2;
      continue;
    }

    if (ch === "_" || ch === "^") {
      const marker = ch;
      i += 1;

      if (i < input.length && input[i] === "(") {
        let depth = 0;
        const start = i;
        while (i < input.length) {
          const groupCh = input[i];
          if (groupCh === "(") {
            depth += 1;
          } else if (groupCh === ")") {
            depth -= 1;
            if (depth === 0) {
              i += 1;
              break;
            }
          }
          i += 1;
        }
        tokens.push(`${marker}${input.slice(start, i)}`);
        continue;
      }

      const start = i;
      while (i < input.length && /[\p{L}0-9]/u.test(input[i])) {
        i += 1;
      }
      if (i > start) {
        tokens.push(`${marker}${input.slice(start, i)}`);
      } else {
        tokens.push(marker);
      }
      continue;
    }

    if (isAsciiLetter(ch)) {
      let j = i + 1;
      while (j < input.length && isAsciiLetter(input[j])) {
        j += 1;
      }
      tokens.push(input.slice(i, j));
      i = j;
      continue;
    }

    if (/[0-9]/.test(ch)) {
      let j = i + 1;
      while (j < input.length && /[0-9.]/.test(input[j])) {
        j += 1;
      }
      tokens.push(input.slice(i, j));
      i = j;
      continue;
    }

    tokens.push(ch);
    i += 1;
  }

  return tokens;
}

function insertImplicitMultiplication(input) {
  const functions = new Set(["sqrt", "sin", "cos", "tan", "cot", "sec", "csc", "log", "ln", "exp", "abs", "min", "max"]);
  const knownIdentifiers = new Set([
    ...functions,
    "alpha",
    "beta",
    "gamma",
    "delta",
    "epsilon",
    "theta",
    "lambda",
    "mu",
    "pi",
    "rho",
    "sigma",
    "tau",
    "phi",
    "omega"
  ]);
  const baseTokens = tokenizeLinearExpression(input);
  const tokens = [];

  for (const token of baseTokens) {
    if (!isIdentifierToken(token)) {
      tokens.push(token);
      continue;
    }

    const lower = token.toLowerCase();
    if (token.length === 1 || knownIdentifiers.has(lower)) {
      tokens.push(token);
      continue;
    }

    const caseParts = token.match(/[A-Z]?[a-z]+|[A-Z]+/g) || [token];
    for (const part of caseParts) {
      const partLower = part.toLowerCase();
      if (part.length === 1 || knownIdentifiers.has(partLower)) {
        tokens.push(part);
      } else {
        for (const ch of part) {
          tokens.push(ch);
        }
      }
    }
  }

  const out = [];

  const needsMulBetween = (prev, curr) => {
    if (!prev || !curr) {
      return false;
    }
    if (prev === "(" || curr === ")" || prev === "," || curr === ",") {
      return false;
    }
    if (prev === "_" || curr === "_") {
      return false;
    }
    if (isScriptToken(prev) || isScriptToken(curr)) {
      return false;
    }
    if (isOperatorToken(prev) || isOperatorToken(curr)) {
      return false;
    }

    if (curr === "(" && isIdentifierToken(prev) && functions.has(prev)) {
      return false;
    }

    const prevEndsOperand = isIdentifierToken(prev) || isNumberToken(prev) || prev === ")";
    const currStartsOperand = isIdentifierToken(curr) || isNumberToken(curr) || curr === "(";
    return prevEndsOperand && currStartsOperand;
  };

  for (const token of tokens) {
    const prev = out.length > 0 ? out[out.length - 1] : null;
    if (needsMulBetween(prev, token)) {
      out.push("*");
    }
    out.push(token);
  }

  return out.join("");
}

function convertLatexToSowisoLinear(input) {
  const greekMap = {
    "\\alpha": "α",
    "\\beta": "β",
    "\\gamma": "γ",
    "\\delta": "δ",
    "\\epsilon": "ε",
    "\\theta": "θ",
    "\\lambda": "λ",
    "\\mu": "μ",
    "\\pi": "π",
    "\\rho": "ρ",
    "\\sigma": "σ",
    "\\tau": "τ",
    "\\phi": "φ",
    "\\omega": "ω"
  };

  let out = input || "";
  out = out.replace(/^\s*\$+|\$+\s*$/g, "");
  out = out.replace(/\\left|\\right|\\,/g, "");
  out = out.replace(/\\(?:hat|widehat|vec|overrightarrow|bar|overline|underline)\s*\{([^{}]+)\}/g, "$1");
  out = out.replace(/\\cdot|\\times/g, "*");
  out = out.replace(/\\leq/g, "<=");
  out = out.replace(/\\geq/g, ">=");
  out = out.replace(/\\neq/g, "!=");
  out = out.replace(/\\pm/g, "+-");

  for (const [latex, plain] of Object.entries(greekMap)) {
    out = out.split(latex).join(plain);
  }

  out = replaceFracLatex(out);
  out = replaceSqrtLatex(out);

  // Normalize powers/subscripts: keep simple indices/exponents compact, group only complex ones.
  out = out.replace(/\^\{([A-Za-z0-9]+)\}/g, "^$1");
  out = out.replace(/_\{([A-Za-z0-9]+)\}/g, "_$1");
  out = out.replace(/\^\{([^{}]+)\}/g, "^($1)");
  out = out.replace(/_\{([^{}]+)\}/g, "_($1)");

  out = out.replace(/[{}]/g, (ch) => (ch === "{" ? "(" : ")"));
  out = out.replace(/\\([a-zA-Z]+)/g, "$1");
  out = out.replace(/\s+/g, " ").trim();
  out = insertImplicitMultiplication(out);
  out = out.replace(/\s+/g, "");
  out = out.replace(/\*{2,}/g, "*");
  out = out.replace(/\(\*/g, "(").replace(/\*\)/g, ")");
  return out;
}

function showEmptyPreview() {
  previewButton.classList.remove("has-image");
  previewImage.removeAttribute("src");
  previewPlaceholder.textContent = "Type a formula to preview it here.";
}

function updatePreview() {
  const raw = latexInput.value.trim();

  if (!raw) {
    showEmptyPreview();
    return;
  }

  const src = `https://latex.codecogs.com/svg.image?${encodeURIComponent(`\\dpi{170} ${raw}`)}`;
  previewImage.src = src;
}

function isInjectableTabUrl(url) {
  return typeof url === "string" && /^(https?|file):\/\//i.test(url);
}

function isSowisoTabUrl(url) {
  return typeof url === "string" && /^https?:\/\/(?:[^/]+\.)?sowiso\.nl\//i.test(url);
}

function getActiveTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      resolve(tabs[0] ?? null);
    });
  });
}

function getFrameIds(tabId) {
  return new Promise((resolve) => {
    if (!chrome.webNavigation || typeof chrome.webNavigation.getAllFrames !== "function") {
      resolve([0]);
      return;
    }

    chrome.webNavigation.getAllFrames({ tabId }, (frames) => {
      if (chrome.runtime.lastError || !Array.isArray(frames) || frames.length === 0) {
        resolve([0]);
        return;
      }

      const frameIds = [...new Set(frames.map((frame) => frame.frameId))];
      resolve(frameIds);
    });
  });
}

function getFrameDetails(tabId) {
  return new Promise((resolve) => {
    if (!chrome.webNavigation || typeof chrome.webNavigation.getAllFrames !== "function") {
      resolve([{ frameId: 0, url: "unknown" }]);
      return;
    }

    chrome.webNavigation.getAllFrames({ tabId }, (frames) => {
      if (chrome.runtime.lastError || !Array.isArray(frames) || frames.length === 0) {
        resolve([{ frameId: 0, url: "unknown" }]);
        return;
      }

      resolve(
        frames.map((frame) => ({
          frameId: frame.frameId,
          parentFrameId: frame.parentFrameId,
          url: frame.url || "unknown"
        }))
      );
    });
  });
}

function focusSowisoSlot(tabId) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        func: () => {
          function isVisible(el) {
            if (!(el instanceof HTMLElement)) {
              return false;
            }
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
          }

          function resolveMarkedSlot() {
            const markedNodes = [...document.querySelectorAll("[data-sowiso-helper-last-slot='1']")]
              .filter((node) => node instanceof HTMLElement && isVisible(node));
            if (markedNodes.length === 0) {
              return null;
            }

            const markedCanvas = markedNodes.find((node) => node instanceof HTMLCanvasElement && node.classList.contains("mathdoxformula"));
            const markedAnchor = markedNodes.find((node) => node instanceof HTMLElement && node.hasAttribute("tabindex"));
            const marked = markedCanvas || markedAnchor || markedNodes[0];
            if (!(marked instanceof HTMLElement)) {
              return null;
            }

            if (marked instanceof HTMLCanvasElement) {
              const anchor = marked.closest("a[tabindex], [tabindex]");
              if (anchor instanceof HTMLElement && isVisible(anchor)) {
                return anchor;
              }
              return marked;
            }

            const td = marked.closest("td");
            if (td) {
              const anchor = td.querySelector("[tabindex]");
              if (anchor instanceof HTMLElement && isVisible(anchor)) {
                return anchor;
              }

              const canvas = td.querySelector("canvas.mathdoxformula");
              if (canvas instanceof HTMLCanvasElement && isVisible(canvas)) {
                return canvas;
              }
            }

            return marked;
          }

          const selectors = [
            "td[class*='pre_input'] + td [tabindex]",
            "td[class*='pre_input'] + td canvas.mathdoxformula",
            "td[class*='pre_input'] + td",
            "table.input_table td.pre_input_text + td [tabindex]",
            "table.input_table td.pre_input_text + td canvas.mathdoxformula",
            "table.input_table td.pre_input_text + td",
            "a[tabindex]",
            "canvas.mathdoxformula"
          ];

          const markedSlot = resolveMarkedSlot();
          if (markedSlot) {
            const events = ["pointerdown", "mousedown", "mouseup", "click"];
            for (const type of events) {
              markedSlot.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            }
            if (markedSlot instanceof HTMLElement) {
              markedSlot.focus();
            }

            return {
              ok: true,
              selector: "[data-sowiso-helper-last-slot='1']",
              tag: markedSlot.tagName,
              className: markedSlot.className || null
            };
          }

          for (const selector of selectors) {
            const nodes = [...document.querySelectorAll(selector)].filter((el) => isVisible(el));
            if (nodes.length === 0) {
              continue;
            }

            let target = nodes[0];
            if (target instanceof HTMLCanvasElement) {
              const anchor = target.closest("a[tabindex], [tabindex]");
              if (anchor instanceof HTMLElement) {
                target = anchor;
              }
            }
            const events = ["pointerdown", "mousedown", "mouseup", "click"];
            for (const type of events) {
              target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            }
            if (target instanceof HTMLElement) {
              target.focus();
            }

            return {
              ok: true,
              selector,
              tag: target.tagName,
              className: target.className || null
            };
          }

          return { ok: false, error: "No visible Sowiso answer slot found in this frame." };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const debugFrames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          result: entry.result || null
        }));

        const successful = (results || []).find((entry) => entry && entry.result && entry.result.ok);
        if (successful) {
          resolve({ ok: true, debugFrames, ...successful.result });
          return;
        }

        const firstError = debugFrames
          .map((entry) => entry.result && entry.result.error)
          .find(Boolean);

        resolve({ ok: false, error: firstError || "Could not focus any Sowiso slot.", debugFrames });
      }
    );
  });
}

function captureSowisoState(tabId) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        func: () => {
          const textareas = [...document.querySelectorAll("textarea.math-editor, textarea.mathdoxformula, textarea[class*='mathdox'], textarea[class*='math-editor']")];
          const values = textareas.map((ta) => ta.value || "");
          const active = document.activeElement;
          return {
            ok: true,
            serial: values.join("||"),
            valueLengths: values.map((v) => v.length),
            activeTag: active ? active.tagName : null,
            activeClass: active && active.className ? String(active.className) : null
          };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const frames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          ...(entry.result || { ok: false })
        }));
        resolve({ ok: true, frames });
      }
    );
  });
}

function waitMs(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function keyDescriptor(ch) {
  const isUpper = /^[A-Z]$/.test(ch);
  const lower = ch.toLowerCase();

  if (/^[a-z]$/i.test(ch)) {
    const keyCode = lower.charCodeAt(0) - 32;
    return { key: ch, code: `Key${lower.toUpperCase()}`, keyCode, modifiers: isUpper ? 8 : 0, text: ch };
  }

  if (/^[0-9]$/.test(ch)) {
    const keyCode = ch.charCodeAt(0);
    return { key: ch, code: `Digit${ch}`, keyCode, modifiers: 0, text: ch };
  }

  const map = {
    " ": { key: " ", code: "Space", keyCode: 32, modifiers: 0, text: " " },
    "=": { key: "=", code: "Equal", keyCode: 187, modifiers: 0, text: "=" },
    "+": { key: "+", code: "Equal", keyCode: 187, modifiers: 8, text: "+" },
    "-": { key: "-", code: "Minus", keyCode: 189, modifiers: 0, text: "-" },
    "_": { key: "_", code: "Minus", keyCode: 189, modifiers: 8, text: "_" },
    "/": { key: "/", code: "Slash", keyCode: 191, modifiers: 0, text: "/" },
    "*": { key: "*", code: "Digit8", keyCode: 56, modifiers: 8, text: "*" },
    "^": { key: "^", code: "Digit6", keyCode: 54, modifiers: 8, text: "^" },
    "(": { key: "(", code: "Digit9", keyCode: 57, modifiers: 8, text: "(" },
    ")": { key: ")", code: "Digit0", keyCode: 48, modifiers: 8, text: ")" },
    "<": { key: "<", code: "Comma", keyCode: 188, modifiers: 8, text: "<" },
    ">": { key: ">", code: "Period", keyCode: 190, modifiers: 8, text: ">" },
    ".": { key: ".", code: "Period", keyCode: 190, modifiers: 0, text: "." },
    ",": { key: ",", code: "Comma", keyCode: 188, modifiers: 0, text: "," }
  };

  if (map[ch]) {
    return map[ch];
  }

  return { key: ch, code: "Unidentified", keyCode: ch.charCodeAt(0) || 0, modifiers: 0, text: ch };
}

function debuggerSend(tabId, method, params) {
  return new Promise((resolve, reject) => {
    chrome.debugger.sendCommand({ tabId }, method, params, () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve();
    });
  });
}

function debuggerAttach(tabId) {
  return new Promise((resolve, reject) => {
    chrome.debugger.attach({ tabId }, "1.3", () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve();
    });
  });
}

function debuggerDetach(tabId) {
  return new Promise((resolve) => {
    chrome.debugger.detach({ tabId }, () => {
      resolve();
    });
  });
}

async function debuggerTypeInsert(tabId, text) {
  const beforeState = await captureSowisoState(tabId);
  const trace = [];
  let attached = false;

  try {
    await debuggerAttach(tabId);
    attached = true;

    for (const ch of text) {
      const descriptor = keyDescriptor(ch);
      await debuggerSend(tabId, "Input.dispatchKeyEvent", {
        type: "rawKeyDown",
        key: descriptor.key,
        code: descriptor.code,
        windowsVirtualKeyCode: descriptor.keyCode,
        nativeVirtualKeyCode: descriptor.keyCode,
        modifiers: descriptor.modifiers
      });

      await debuggerSend(tabId, "Input.dispatchKeyEvent", {
        type: "char",
        key: descriptor.key,
        code: descriptor.code,
        text: descriptor.text,
        unmodifiedText: descriptor.text,
        windowsVirtualKeyCode: descriptor.keyCode,
        nativeVirtualKeyCode: descriptor.keyCode,
        modifiers: descriptor.modifiers
      });

      await debuggerSend(tabId, "Input.dispatchKeyEvent", {
        type: "keyUp",
        key: descriptor.key,
        code: descriptor.code,
        windowsVirtualKeyCode: descriptor.keyCode,
        nativeVirtualKeyCode: descriptor.keyCode,
        modifiers: descriptor.modifiers
      });

      trace.push({ char: ch, code: descriptor.code, keyCode: descriptor.keyCode, modifiers: descriptor.modifiers });
      await waitMs(14);
    }
  } catch (error) {
    if (attached) {
      await debuggerDetach(tabId);
    }
    return { ok: false, error: error.message, trace };
  }

  if (attached) {
    await debuggerDetach(tabId);
  }

  const afterState = await captureSowisoState(tabId);
  const beforeSerial = beforeState.ok ? beforeState.frames.map((frame) => frame.serial || "").join("||") : "";
  const afterSerial = afterState.ok ? afterState.frames.map((frame) => frame.serial || "").join("||") : "";
  const changed = beforeSerial !== afterSerial;

  return {
    ok: changed,
    changed,
    trace,
    beforeState: beforeState.ok ? beforeState.frames : [],
    afterState: afterState.ok ? afterState.frames : [],
    error: changed ? undefined : "Debugger typing sent keys, but editor state did not change."
  };
}

function sowisoTextareaInsert(tabId, formula) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        args: [formula],
        func: (rawFormula) => {
          function isVisible(el) {
            if (!(el instanceof HTMLElement)) {
              return false;
            }
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
          }

          function clickLikeUser(el) {
            const events = ["pointerdown", "mousedown", "mouseup", "click"];
            for (const type of events) {
              el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            }
          }

          function resolveMarkedSlot() {
            const markedNodes = [...document.querySelectorAll("[data-sowiso-helper-last-slot='1']")]
              .filter((node) => node instanceof HTMLElement && isVisible(node));
            if (markedNodes.length === 0) {
              return null;
            }

            const markedCanvas = markedNodes.find((node) => node instanceof HTMLCanvasElement && node.classList.contains("mathdoxformula"));
            const markedAnchor = markedNodes.find((node) => node instanceof HTMLElement && node.hasAttribute("tabindex"));
            const marked = markedCanvas || markedAnchor || markedNodes[0];
            if (!(marked instanceof HTMLElement)) {
              return null;
            }

            if (marked instanceof HTMLCanvasElement) {
              const anchor = marked.closest("a[tabindex], [tabindex]");
              if (anchor instanceof HTMLElement && isVisible(anchor)) {
                return anchor;
              }
              return marked;
            }

            const td = marked.closest("td");
            if (td) {
              const anchor = td.querySelector("[tabindex]");
              if (anchor instanceof HTMLElement && isVisible(anchor)) {
                return anchor;
              }
              const canvas = td.querySelector("canvas.mathdoxformula");
              if (canvas instanceof HTMLCanvasElement && isVisible(canvas)) {
                return canvas;
              }
            }

            return marked;
          }

          function findBundleForActiveSlot() {
            const active = document.activeElement;
            if (active instanceof HTMLElement && (active.classList.contains("mathdoxformula") || active.hasAttribute("tabindex"))) {
              const td = active.closest("td");
              if (td) {
                const ta = td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']");
                const canvas = td.querySelector("canvas.mathdoxformula");
                if (ta instanceof HTMLTextAreaElement) {
                  return {
                    textarea: ta,
                    canvas: canvas instanceof HTMLCanvasElement ? canvas : null
                  };
                }
              }
            }

            const marked = resolveMarkedSlot();
            if (marked) {
              const td = marked.closest("td");
              if (td) {
                const ta = td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']");
                const canvas = td.querySelector("canvas.mathdoxformula");
                if (ta instanceof HTMLTextAreaElement) {
                  clickLikeUser(marked);
                  marked.focus();
                  return {
                    textarea: ta,
                    canvas: canvas instanceof HTMLCanvasElement ? canvas : null
                  };
                }
              }
            }

            const slot = [...document.querySelectorAll("td[class*='pre_input'] + td [tabindex], td[class*='pre_input'] + td canvas.mathdoxformula, table.input_table td.pre_input_text + td [tabindex], table.input_table td.pre_input_text + td canvas.mathdoxformula, a[tabindex], canvas.mathdoxformula")]
              .find((el) => isVisible(el));
            if (slot) {
              let target = slot;
              if (target instanceof HTMLCanvasElement) {
                const anchor = target.closest("a[tabindex], [tabindex]");
                if (anchor instanceof HTMLElement) {
                  target = anchor;
                }
              }

              clickLikeUser(target);
              if (target instanceof HTMLElement) {
                target.focus();
              }
              const td = target.closest("td") || slot.closest("td");
              if (td) {
                const ta = td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']");
                const canvas = td.querySelector("canvas.mathdoxformula");
                if (ta instanceof HTMLTextAreaElement) {
                  return {
                    textarea: ta,
                    canvas: canvas instanceof HTMLCanvasElement ? canvas : null
                  };
                }
              }
            }

            return null;
          }

          function readCanvasSignature(canvas) {
            if (!(canvas instanceof HTMLCanvasElement) || typeof canvas.toDataURL !== "function") {
              return null;
            }
            try {
              return canvas.toDataURL();
            } catch (_error) {
              return null;
            }
          }

          function readVisibleCanvasSignature() {
            const canvases = [...document.querySelectorAll("td[class*='pre_input'] + td canvas.mathdoxformula, table.input_table td.pre_input_text + td canvas.mathdoxformula, canvas.mathdoxformula")]
              .filter((el) => el instanceof HTMLCanvasElement && isVisible(el));
            const serial = canvases.map((canvas) => readCanvasSignature(canvas)).join("||");
            return { count: canvases.length, serial };
          }

          const bundle = findBundleForActiveSlot();
          const target = bundle ? bundle.textarea : null;
          const targetCanvas = bundle ? bundle.canvas : null;
          if (!target) {
            return { ok: false, error: "Could not find MathDox textarea for active answer slot." };
          }

          const before = target.value || "";
          const beforeCanvas = readCanvasSignature(targetCanvas);
          const beforeVisibleCanvas = readVisibleCanvasSignature();
          target.focus();

          // Select all existing content so execCommand replaces it.
          if (typeof target.setSelectionRange === "function") {
            target.setSelectionRange(0, target.value.length);
          } else {
            target.select();
          }

          // execCommand('insertText') generates trusted InputEvents (isTrusted: true),
          // which MathDox requires. Untrusted events from dispatchEvent are ignored.
          let execCommandWorked = false;
          try {
            execCommandWorked = document.execCommand("insertText", false, rawFormula);
          } catch (_error) {
            execCommandWorked = false;
          }

          // Fallback: set value directly if execCommand didn't work.
          if (!execCommandWorked || target.value !== rawFormula) {
            if (typeof target.setRangeText === "function") {
              target.setRangeText(rawFormula, 0, target.value.length, "end");
            } else {
              target.value = rawFormula;
            }

            target.dispatchEvent(new Event("input", { bubbles: true }));
            target.dispatchEvent(new Event("change", { bubbles: true }));
          }

          const after = target.value || "";
          const afterCanvas = readCanvasSignature(targetCanvas);
          const afterVisibleCanvas = readVisibleCanvasSignature();
          const valueChanged = after !== before;
          const canvasChanged = beforeCanvas !== null && afterCanvas !== null && beforeCanvas !== afterCanvas;
          const anyVisibleCanvasChanged = beforeVisibleCanvas.serial !== afterVisibleCanvas.serial;
          const hasVisibleCanvas = afterVisibleCanvas.count > 0 || beforeVisibleCanvas.count > 0 || Boolean(targetCanvas);
          const visibleEffect = canvasChanged || anyVisibleCanvasChanged;
          const ok = hasVisibleCanvas ? visibleEffect : (after === rawFormula || valueChanged);

          return {
            ok,
            execCommandWorked,
            valueChanged,
            canvasChanged,
            anyVisibleCanvasChanged,
            hasVisibleCanvas,
            beforeLength: before.length,
            afterLength: after.length,
            className: target.className || null
          };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const debugFrames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          result: entry.result || null
        }));

        const successful = (results || []).find((entry) => entry && entry.result && entry.result.ok);
        if (successful) {
          resolve({ ok: true, debugFrames, ...successful.result });
          return;
        }

        const firstError = debugFrames
          .map((entry) => entry.result && entry.result.error)
          .find(Boolean);

        resolve({ ok: false, error: firstError || "MathDox textarea insertion failed.", debugFrames });
      }
    );
  });
}

function sowisoSlotKeyTypeInsert(tabId, formula) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        args: [formula],
        func: async (formulaText) => {
          function wait(ms) {
            return new Promise((done) => window.setTimeout(done, ms));
          }

          function isVisible(el) {
            if (!(el instanceof HTMLElement)) {
              return false;
            }
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
          }

          function normalizeSlot(el) {
            if (!(el instanceof HTMLElement)) {
              return null;
            }
            if (el instanceof HTMLCanvasElement) {
              const anchor = el.closest("a[tabindex], [tabindex]");
              if (anchor instanceof HTMLElement) {
                return anchor;
              }
            }
            return el;
          }

          function resolveMarkedSlot() {
            const markedNodes = [...document.querySelectorAll("[data-sowiso-helper-last-slot='1']")]
              .filter((node) => node instanceof HTMLElement && isVisible(node));
            if (markedNodes.length === 0) {
              return null;
            }
            const markedCanvas = markedNodes.find((node) => node instanceof HTMLCanvasElement && node.classList.contains("mathdoxformula"));
            const markedAnchor = markedNodes.find((node) => node instanceof HTMLElement && node.hasAttribute("tabindex"));
            const marked = markedCanvas || markedAnchor || markedNodes[0];
            return normalizeSlot(marked);
          }

          function findSlot() {
            const marked = resolveMarkedSlot();
            if (marked) {
              return marked;
            }

            const active = document.activeElement;
            if (active instanceof HTMLElement) {
              const normalized = normalizeSlot(active);
              if (normalized && (normalized.classList.contains("mathdoxformula") || normalized.hasAttribute("tabindex"))) {
                return normalized;
              }
            }

            const selectors = [
              "td[class*='pre_input'] + td [tabindex]",
              "td[class*='pre_input'] + td canvas.mathdoxformula",
              "table.input_table td.pre_input_text + td [tabindex]",
              "table.input_table td.pre_input_text + td canvas.mathdoxformula",
              "a[tabindex]",
              "canvas.mathdoxformula"
            ];
            for (const selector of selectors) {
              const candidate = [...document.querySelectorAll(selector)].find((node) => isVisible(node));
              if (candidate) {
                return normalizeSlot(candidate);
              }
            }
            return null;
          }

          function snapshot() {
            const textareas = [...document.querySelectorAll("textarea.math-editor, textarea.mathdoxformula, textarea[class*='mathdox'], textarea[class*='math-editor']")];
            const serial = textareas.map((ta) => ta.value || "").join("||");
            const canvases = [...document.querySelectorAll("td[class*='pre_input'] + td canvas.mathdoxformula, table.input_table td.pre_input_text + td canvas.mathdoxformula, canvas.mathdoxformula")]
              .filter((canvas) => canvas instanceof HTMLCanvasElement && isVisible(canvas));
            const canvasSerial = canvases.map((canvas) => {
              try {
                return canvas.toDataURL();
              } catch (_error) {
                return null;
              }
            }).join("||");
            return { serial, canvasSerial };
          }

          function codeForChar(ch) {
            if (/^[a-z]$/i.test(ch)) return `Key${ch.toUpperCase()}`;
            if (/^[0-9]$/.test(ch)) return `Digit${ch}`;
            const map = {
              "=": "Equal",
              "+": "Equal",
              "-": "Minus",
              "_": "Minus",
              "/": "Slash",
              "?": "Slash",
              ".": "Period",
              ",": "Comma",
              "(": "Digit9",
              ")": "Digit0",
              "^": "Digit6",
              "*": "Digit8",
              "<": "Comma",
              ">": "Period"
            };
            return map[ch] || "Unidentified";
          }

          const slot = findSlot();
          if (!slot) {
            return { ok: false, error: "No active MathDox slot found for direct key typing." };
          }

          const td = slot.closest("td");
          const ta = td ? td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']") : null;
          const targetTextarea = ta instanceof HTMLTextAreaElement ? ta : null;

          const before = snapshot();

          const clickTypes = ["pointerdown", "mousedown", "mouseup", "click"];
          for (const type of clickTypes) {
            slot.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
          }
          slot.focus();
          await wait(25);

          if (targetTextarea) {
            targetTextarea.focus();
            if (typeof targetTextarea.setRangeText === "function") {
              targetTextarea.setRangeText("", 0, targetTextarea.value.length, "end");
            } else {
              targetTextarea.value = "";
            }
            targetTextarea.dispatchEvent(new Event("input", { bubbles: true }));
            targetTextarea.dispatchEvent(new Event("change", { bubbles: true }));
          }

          const fire = (receiver, type, payload) => {
            receiver.dispatchEvent(
              new KeyboardEvent(type, {
                bubbles: true,
                cancelable: true,
                composed: true,
                ...payload
              })
            );
          };

          for (const ch of [...(formulaText || "")]) {
            const payload = {
              key: ch,
              code: codeForChar(ch),
              shiftKey: /^[A-Z]$/.test(ch)
            };

            fire(slot, "keydown", payload);
            fire(document, "keydown", payload);
            fire(slot, "keypress", payload);
            fire(document, "keypress", payload);

            if (targetTextarea) {
              const end = targetTextarea.value.length;
              if (typeof targetTextarea.setRangeText === "function") {
                targetTextarea.setRangeText(ch, end, end, "end");
              } else {
                targetTextarea.value = `${targetTextarea.value}${ch}`;
              }
              try {
                targetTextarea.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: ch, inputType: "insertText" }));
                targetTextarea.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: false, data: ch, inputType: "insertText" }));
              } catch (_error) {
                targetTextarea.dispatchEvent(new Event("input", { bubbles: true }));
              }
            }

            fire(slot, "keyup", payload);
            fire(document, "keyup", payload);
          }

          const enterPayload = { key: "Enter", code: "Enter", shiftKey: false };
          fire(slot, "keydown", enterPayload);
          fire(document, "keydown", enterPayload);
          fire(slot, "keyup", enterPayload);
          fire(document, "keyup", enterPayload);

          await wait(140);
          const after = snapshot();
          const valueChanged = before.serial !== after.serial;
          const canvasChanged = before.canvasSerial !== after.canvasSerial;

          return {
            ok: canvasChanged || valueChanged,
            valueChanged,
            canvasChanged,
            textareaPresent: Boolean(targetTextarea)
          };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const debugFrames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          result: entry.result || null
        }));

        const successful = (results || []).find((entry) => entry && entry.result && entry.result.ok);
        if (successful) {
          resolve({ ok: true, debugFrames, ...successful.result });
          return;
        }

        const firstError = debugFrames
          .map((entry) => entry.result && entry.result.error)
          .find(Boolean);

        resolve({ ok: false, error: firstError || "Direct slot key typing failed.", debugFrames });
      }
    );
  });
}

function sowisoMainWorldApiInsert(tabId, formula, latexSource) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        world: "MAIN",
        args: [formula, latexSource],
        func: async (rawFormula, rawLatexSource) => {
          function isVisible(el) {
            if (!(el instanceof HTMLElement)) {
              return false;
            }
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
          }

          function readCanvasSignature(canvas) {
            if (!(canvas instanceof HTMLCanvasElement) || typeof canvas.toDataURL !== "function") {
              return null;
            }
            try {
              return canvas.toDataURL();
            } catch (_error) {
              return null;
            }
          }

          function readVisibleCanvasState() {
            const canvases = [...document.querySelectorAll("td[class*='pre_input'] + td canvas.mathdoxformula, table.input_table td.pre_input_text + td canvas.mathdoxformula, canvas.mathdoxformula")]
              .filter((canvas) => canvas instanceof HTMLCanvasElement && isVisible(canvas));
            return {
              count: canvases.length,
              serial: canvases.map((canvas) => readCanvasSignature(canvas)).join("||")
            };
          }

          function normalizeSlot(el) {
            if (!(el instanceof HTMLElement)) {
              return null;
            }
            if (el instanceof HTMLCanvasElement) {
              const anchor = el.closest("a[tabindex], [tabindex]");
              if (anchor instanceof HTMLElement) {
                return anchor;
              }
            }
            return el;
          }

          function resolveMarkedSlot() {
            const markedNodes = [...document.querySelectorAll("[data-sowiso-helper-last-slot='1']")]
              .filter((node) => node instanceof HTMLElement && isVisible(node));
            if (markedNodes.length === 0) {
              return null;
            }
            const markedCanvas = markedNodes.find((node) => node instanceof HTMLCanvasElement && node.classList.contains("mathdoxformula"));
            const markedAnchor = markedNodes.find((node) => node instanceof HTMLElement && node.hasAttribute("tabindex"));
            const marked = markedCanvas || markedAnchor || markedNodes[0];
            return normalizeSlot(marked);
          }

          function findSlotAndTextarea() {
            const marked = resolveMarkedSlot();
            const active = normalizeSlot(document.activeElement);
            const selectors = [
              "td[class*='pre_input'] + td [tabindex]",
              "td[class*='pre_input'] + td canvas.mathdoxformula",
              "table.input_table td.pre_input_text + td [tabindex]",
              "table.input_table td.pre_input_text + td canvas.mathdoxformula",
              "a[tabindex]",
              "canvas.mathdoxformula"
            ];

            const candidates = [];
            if (marked && isVisible(marked)) {
              candidates.push(marked);
            }
            if (active && isVisible(active)) {
              candidates.push(active);
            }
            for (const selector of selectors) {
              const node = [...document.querySelectorAll(selector)].find((candidate) => isVisible(candidate));
              if (node) {
                candidates.push(normalizeSlot(node));
              }
            }

            for (const candidate of candidates) {
              const slot = normalizeSlot(candidate);
              if (!slot) {
                continue;
              }

              const td = slot.closest("td");
              if (!td) {
                continue;
              }

              const textarea = td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']");
              const canvas = td.querySelector("canvas.mathdoxformula");
              if (!(textarea instanceof HTMLTextAreaElement)) {
                continue;
              }

              return {
                slot,
                textarea,
                canvas: canvas instanceof HTMLCanvasElement ? canvas : null
              };
            }

            return { slot: null, textarea: null, canvas: null };
          }

          function methodList(obj) {
            if (!obj) {
              return [];
            }
            const methods = new Set();
            let cur = obj;
            let depth = 0;
            while (cur && depth < 5) {
              for (const key of Object.getOwnPropertyNames(cur)) {
                try {
                  if (typeof obj[key] === "function") {
                    methods.add(key);
                  }
                } catch (_error) {
                  // Ignore access errors.
                }
              }
              cur = Object.getPrototypeOf(cur);
              depth += 1;
            }
            return [...methods];
          }

          function readState(bundle) {
            const value = bundle.textarea ? (bundle.textarea.value || "") : "";
            const canvas = readCanvasSignature(bundle.canvas);
            const visible = readVisibleCanvasState();
            return { value, canvas, visible };
          }

          const bundle = findSlotAndTextarea();
          if (!bundle.slot) {
            return { ok: false, error: "No active slot found in MAIN world." };
          }

          const clicks = ["pointerdown", "mousedown", "mouseup", "click"];
          for (const type of clicks) {
            bundle.slot.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
          }
          bundle.slot.focus();

          const beforeState = readState(bundle);

          const attempts = [];
          const attemptLimit = 240;
          const pushAttempt = (entry) => {
            if (attempts.length < attemptLimit) {
              attempts.push(entry);
            }
          };

          const potentialEditors = [];
          const seenEditors = new Set();
          const addEditor = (source, editor) => {
            if (!editor || typeof editor !== "object") {
              return;
            }
            if (seenEditors.has(editor)) {
              return;
            }
            seenEditors.add(editor);
            potentialEditors.push({ source, editor });
          };

          if (bundle.textarea && bundle.textarea.formulaeditorobject) {
            addEditor("textarea.formulaeditorobject", bundle.textarea.formulaeditorobject);
          }
          if (bundle.canvas && bundle.canvas.formulaeditorobject) {
            addEditor("canvas.formulaeditorobject", bundle.canvas.formulaeditorobject);
          }
          if (bundle.slot && bundle.slot.formulaeditorobject) {
            addEditor("slot.formulaeditorobject", bundle.slot.formulaeditorobject);
          }

          const harvestEditorsFromHost = (host, source) => {
            if (!host || (typeof host !== "object" && typeof host !== "function")) {
              return;
            }
            for (const key of Object.getOwnPropertyNames(host)) {
              if (!/(editor|formula|mathdox|math)/i.test(key)) {
                continue;
              }
              let value;
              try {
                value = host[key];
              } catch (_error) {
                continue;
              }
              if (value && (typeof value === "object" || typeof value === "function")) {
                addEditor(`${source}.${key}`, value);
              }
            }
          };

          harvestEditorsFromHost(bundle.slot, "slot");
          harvestEditorsFromHost(bundle.canvas, "canvas");
          harvestEditorsFromHost(bundle.textarea, "textarea");

          const currentValue = () => (bundle.textarea ? (bundle.textarea.value || "") : "");
          const setTextareaValue = (value, source) => {
            if (!bundle.textarea) {
              return false;
            }
            const before = currentValue();
            try {
              bundle.textarea.value = "";
              bundle.textarea.dispatchEvent(new Event("input", { bubbles: true }));
              bundle.textarea.dispatchEvent(new Event("change", { bubbles: true }));
              bundle.textarea.value = value;
              bundle.textarea.dispatchEvent(new Event("input", { bubbles: true }));
              bundle.textarea.dispatchEvent(new Event("change", { bubbles: true }));
              pushAttempt({ source, method: "textarea.value-clear-then-set", ok: true, beforeLength: before.length, afterLength: value.length });
              return true;
            } catch (error) {
              pushAttempt({ source, method: "textarea.value-clear-then-set", ok: false, error: error && error.message ? error.message : String(error) });
              return false;
            }
          };

          const changedFrom = (baseline) => {
            const next = readState(bundle);
            const valueChanged = next.value !== baseline.value;
            const canvasChanged = baseline.canvas !== null && next.canvas !== null && baseline.canvas !== next.canvas;
            const anyVisibleCanvasChanged = baseline.visible.serial !== next.visible.serial;
            const hasVisibleCanvas = baseline.visible.count > 0 || next.visible.count > 0 || Boolean(bundle.canvas);
            const ok = hasVisibleCanvas ? (canvasChanged || anyVisibleCanvasChanged) : valueChanged;
            return {
              ok,
              valueChanged,
              canvasChanged,
              anyVisibleCanvasChanged,
              hasVisibleCanvas,
              next
            };
          };

          const invokeNoArgMethods = (target, source, methods) => {
            if (!target) {
              return false;
            }
            let any = false;
            for (const method of methods) {
              if (typeof target[method] !== "function") {
                continue;
              }
              try {
                target[method]();
                pushAttempt({ source, method, ok: true });
                any = true;
              } catch (error) {
                pushAttempt({ source, method, ok: false, error: error && error.message ? error.message : String(error) });
              }
            }
            return any;
          }

          const orgFormulaEditor = window.org && window.org.mathdox && window.org.mathdox.formulaeditor
            ? window.org.mathdox.formulaeditor
            : null;
          const formulaEditorClass = orgFormulaEditor && orgFormulaEditor.FormulaEditor ? orgFormulaEditor.FormulaEditor : null;

          if (formulaEditorClass) {
            const pickerCalls = [
              { source: "FormulaEditor.getFocusedEditor", fn: "getFocusedEditor", args: [] },
              { source: "FormulaEditor.getLastFocusedEditor", fn: "getLastFocusedEditor", args: [] },
              { source: "FormulaEditor.getEditorByCanvas", fn: "getEditorByCanvas", args: [bundle.canvas] },
              { source: "FormulaEditor.getEditorByCanvas(slot)", fn: "getEditorByCanvas", args: [bundle.slot] },
              { source: "FormulaEditor.getEditorByTextArea", fn: "getEditorByTextArea", args: [bundle.textarea] },
              { source: "FormulaEditor.getEditorByTextArea(id)", fn: "getEditorByTextArea", args: [bundle.textarea ? bundle.textarea.id : null] }
            ];

            for (const call of pickerCalls) {
              const hasNullArgs = call.args.some((arg) => arg === null || arg === undefined || arg === "");
              if (hasNullArgs) {
                continue;
              }
              if (typeof formulaEditorClass[call.fn] !== "function") {
                continue;
              }
              try {
                const editor = formulaEditorClass[call.fn](...call.args);
                if (editor) {
                  pushAttempt({ source: call.source, method: call.fn, ok: true, returnedEditor: true });
                  addEditor(call.source, editor);
                } else {
                  pushAttempt({ source: call.source, method: call.fn, ok: true, returnedEditor: false });
                }
              } catch (error) {
                pushAttempt({ source: call.source, method: call.fn, ok: false, error: error && error.message ? error.message : String(error) });
              }
            }

            const scanForEditors = (host, source) => {
              if (!host || (typeof host !== "object" && typeof host !== "function")) {
                return;
              }
              for (const key of Object.getOwnPropertyNames(host)) {
                if (!/(editor|instance|list|pool|registry|cache)/i.test(key)) {
                  continue;
                }
                let value;
                try {
                  value = host[key];
                } catch (_error) {
                  continue;
                }
                if (Array.isArray(value)) {
                  for (let idx = 0; idx < value.length; idx += 1) {
                    addEditor(`${source}.${key}[${idx}]`, value[idx]);
                  }
                } else if (value && typeof value === "object") {
                  addEditor(`${source}.${key}`, value);
                }
              }
            };

            scanForEditors(formulaEditorClass, "FormulaEditor");
            scanForEditors(orgFormulaEditor, "org.mathdox.formulaeditor");
          }

          const latexCandidateRaw = typeof rawLatexSource === "string" ? rawLatexSource.trim() : "";
          const latexCandidate = latexCandidateRaw
            .replace(/^\s*\$\$?/, "")
            .replace(/\$\$?\s*$/, "")
            .trim();
          let preparedMathML = null;
          const hasLatexCommands = /\\[A-Za-z]+/.test(latexCandidate || rawFormula);

          const wrapMathML = (value) => {
            const trimmed = typeof value === "string" ? value.trim() : "";
            if (!trimmed) {
              return "";
            }
            if (/^<\s*math[\s>]/i.test(trimmed)) {
              return trimmed;
            }
            return `<math xmlns="http://www.w3.org/1998/Math/MathML">${trimmed}</math>`;
          };

          const buildPreparedMathML = (rawMathML, source) => {
            const mathmlString = wrapMathML(rawMathML);
            if (!mathmlString) {
              return null;
            }
            try {
              const parsedXml = new DOMParser().parseFromString(mathmlString, "application/xml");
              const parseError = parsedXml.querySelector("parsererror");
              if (parseError) {
                pushAttempt({
                  source,
                  method: "parseMathML",
                  ok: false,
                  error: parseError.textContent ? parseError.textContent.trim().slice(0, 300) : "XML parsererror"
                });
                return null;
              }

              const root = parsedXml.documentElement;
              if (!root || root.tagName.toLowerCase() !== "math") {
                pushAttempt({ source, method: "parseMathML", ok: false, error: "Parsed MathML root was not <math>." });
                return null;
              }

              const mathNs = "http://www.w3.org/1998/Math/MathML";
              const allowedTags = new Set([
                "math",
                "mrow",
                "mi",
                "mn",
                "mo",
                "mtext",
                "ms",
                "mspace",
                "mfrac",
                "msqrt",
                "mroot",
                "msub",
                "msup",
                "msubsup",
                "mover",
                "munder",
                "munderover",
                "mfenced",
                "mtable",
                "mtr",
                "mtd",
                "mmultiscripts",
                "mprescripts",
                "none"
              ]);
              const tokenTags = new Set(["mi", "mn", "mo", "mtext", "ms"]);
              const allowedAttrs = new Set([
                "stretchy",
                "form",
                "fence",
                "separator",
                "separators",
                "open",
                "close",
                "accent",
                "accentunder",
                "largeop",
                "movablelimits",
                "mathvariant",
                "mathsize",
                "mathcolor"
              ]);

              const sanitizeNode = (node, outDoc) => {
                if (!node) {
                  return null;
                }

                if (node.nodeType === Node.TEXT_NODE) {
                  const text = node.textContent || "";
                  return text.length > 0 ? outDoc.createTextNode(text) : null;
                }

                if (node.nodeType !== Node.ELEMENT_NODE) {
                  return null;
                }

                const tag = (node.localName || node.nodeName || "").toLowerCase();
                if (!tag) {
                  return null;
                }

                if (tag === "annotation" || tag === "annotation-xml") {
                  return null;
                }

                if (tag === "semantics") {
                  const elementChildren = [...node.childNodes].filter((child) => child.nodeType === Node.ELEMENT_NODE);
                  const preferred = elementChildren.find((child) => {
                    const childTag = (child.localName || child.nodeName || "").toLowerCase();
                    return childTag && childTag !== "annotation" && childTag !== "annotation-xml";
                  });
                  if (preferred) {
                    return sanitizeNode(preferred, outDoc);
                  }
                  return null;
                }

                if (tag === "maction") {
                  const elementChildren = [...node.childNodes]
                    .filter((child) => child.nodeType === Node.ELEMENT_NODE);
                  if (elementChildren.length === 0) {
                    return null;
                  }
                  const rawSelection = Number.parseInt(node.getAttribute("selection") || "1", 10);
                  const selection = Number.isFinite(rawSelection) ? rawSelection : 1;
                  const selectedIndex = Math.max(0, Math.min(elementChildren.length - 1, selection - 1));
                  const selected = elementChildren[selectedIndex] || elementChildren[0];
                  return sanitizeNode(selected, outDoc);
                }

                if (tag === "mstyle" || tag === "mpadded" || tag === "mphantom") {
                  const childNodes = [...node.childNodes]
                    .map((child) => sanitizeNode(child, outDoc))
                    .filter(Boolean);
                  if (childNodes.length === 0) {
                    return null;
                  }
                  if (childNodes.length === 1) {
                    return childNodes[0];
                  }
                  const wrapper = outDoc.createElementNS(mathNs, "mrow");
                  for (const childNode of childNodes) {
                    wrapper.appendChild(childNode);
                  }
                  return wrapper;
                }

                if (!allowedTags.has(tag)) {
                  const childNodes = [...node.childNodes]
                    .map((child) => sanitizeNode(child, outDoc))
                    .filter(Boolean);
                  if (childNodes.length === 0) {
                    return null;
                  }
                  if (childNodes.length === 1) {
                    return childNodes[0];
                  }
                  const wrapper = outDoc.createElementNS(mathNs, "mrow");
                  for (const childNode of childNodes) {
                    wrapper.appendChild(childNode);
                  }
                  return wrapper;
                }

                const outEl = outDoc.createElementNS(mathNs, tag);
                if (node.attributes && node.attributes.length > 0) {
                  for (const attr of node.attributes) {
                    const name = (attr && attr.name) ? attr.name.toLowerCase() : "";
                    if (!name || name.startsWith("data-")) {
                      continue;
                    }
                    if (name === "class" || name === "id" || name === "style" || name === "xmlns") {
                      continue;
                    }
                    if (!allowedAttrs.has(name)) {
                      continue;
                    }
                    outEl.setAttribute(name, attr.value);
                  }
                }

                const children = [...node.childNodes]
                  .map((child) => sanitizeNode(child, outDoc))
                  .filter(Boolean);

                if (children.length > 0) {
                  for (const child of children) {
                    outEl.appendChild(child);
                  }
                } else if (tokenTags.has(tag)) {
                  outEl.textContent = node.textContent || "";
                }

                return outEl;
              };

              const outDoc = document.implementation.createDocument(mathNs, "math", null);
              const sanitized = sanitizeNode(root, outDoc);
              if (!sanitized) {
                pushAttempt({ source, method: "sanitizeMathML", ok: false, error: "Sanitizer produced empty output." });
                return null;
              }

              const outRoot = outDoc.documentElement;
              while (outRoot.firstChild) {
                outRoot.removeChild(outRoot.firstChild);
              }

              if (sanitized.nodeType === Node.ELEMENT_NODE && sanitized.localName && sanitized.localName.toLowerCase() === "math") {
                for (const child of [...sanitized.childNodes]) {
                  outRoot.appendChild(child.cloneNode(true));
                }
              } else {
                outRoot.appendChild(sanitized);
              }

              const sanitizedString = new XMLSerializer().serializeToString(outRoot);
              return { string: sanitizedString, node: null };
            } catch (error) {
              pushAttempt({
                source,
                method: "parseMathML",
                ok: false,
                error: error && error.message ? error.message : String(error)
              });
              return null;
            }
          };

          const convertLatexViaMathJaxHub = async (latex) => {
            const mj = window.MathJax;
            const hub = mj && mj.Hub;
            if (!hub || typeof hub.Queue !== "function" || typeof hub.getAllJax !== "function") {
              pushAttempt({
                source: "MathJax.Hub",
                method: "availability",
                ok: false,
                error: "MathJax Hub API unavailable."
              });
              return null;
            }

            if (!(document.body instanceof HTMLElement)) {
              pushAttempt({
                source: "MathJax.Hub",
                method: "availability",
                ok: false,
                error: "Document body unavailable."
              });
              return null;
            }

            return await new Promise((resolveMathML) => {
              const host = document.createElement("span");
              host.style.position = "fixed";
              host.style.left = "-10000px";
              host.style.top = "-10000px";
              host.style.opacity = "0";
              host.style.pointerEvents = "none";

              const script = document.createElement("script");
              script.type = "math/tex";
              script.text = latex;
              host.appendChild(script);
              document.body.appendChild(host);

              let settled = false;
              const settle = (value, meta) => {
                if (settled) {
                  return;
                }
                settled = true;
                window.clearTimeout(timeoutId);
                try {
                  host.remove();
                } catch (_removeError) {
                  // Ignore remove errors.
                }
                resolveMathML(value);
                if (meta) {
                  pushAttempt(meta);
                }
              };

              const timeoutId = window.setTimeout(() => {
                settle(null, {
                  source: "MathJax.Hub",
                  method: "toMathML",
                  ok: false,
                  error: "Timed out waiting for MathJax typeset."
                });
              }, 2000);

              hub.Queue(
                ["Typeset", hub, host],
                () => {
                  try {
                    const allJax = hub.getAllJax(host);
                    const jax = Array.isArray(allJax) && allJax.length > 0 ? allJax[0] : null;
                    if (!jax || !jax.root || typeof jax.root.toMathML !== "function") {
                      settle(null, {
                        source: "MathJax.Hub",
                        method: "toMathML",
                        ok: false,
                        error: "No MathJax Jax root available after typeset."
                      });
                      return;
                    }

                    const rawMathML = jax.root.toMathML("");
                    const wrappedMathML = wrapMathML(rawMathML);
                    settle(wrappedMathML, {
                      source: "MathJax.Hub",
                      method: "toMathML",
                      ok: Boolean(wrappedMathML),
                      latexLength: latex.length
                    });
                  } catch (error) {
                    settle(null, {
                      source: "MathJax.Hub",
                      method: "toMathML",
                      ok: false,
                      error: error && error.message ? error.message : String(error)
                    });
                  }
                }
              );
            });
          };

          if (hasLatexCommands) {
            const hubMathML = await convertLatexViaMathJaxHub(latexCandidate || rawFormula);
            if (hubMathML) {
              preparedMathML = buildPreparedMathML(hubMathML, "MathJax.Hub");
            }
          }

          if (!preparedMathML && hasLatexCommands) {
            try {
              const mj = window.MathJax;
              const TeX = mj && mj.InputJax && mj.InputJax.TeX ? mj.InputJax.TeX : null;
              if (TeX && typeof TeX.Parse === "function") {
                const parsed = TeX.Parse(latexCandidate).mml();
                if (parsed && typeof parsed.toMathML === "function") {
                  let inner = "";
                  let mathmlResultOk = false;
                  try {
                    inner = parsed.toMathML();
                    mathmlResultOk = typeof inner === "string" && inner.trim().length > 0;
                  } catch (primaryError) {
                    pushAttempt({
                      source: "MathJax.TeX.Parse",
                      method: "toMathML()",
                      ok: false,
                      error: primaryError && primaryError.message ? primaryError.message : String(primaryError)
                    });
                  }
                  if (!mathmlResultOk) {
                    try {
                      inner = parsed.toMathML("");
                      mathmlResultOk = typeof inner === "string" && inner.trim().length > 0;
                    } catch (secondaryError) {
                      pushAttempt({
                        source: "MathJax.TeX.Parse",
                        method: "toMathML('')",
                        ok: false,
                        error: secondaryError && secondaryError.message ? secondaryError.message : String(secondaryError)
                      });
                    }
                  }
                  if (mathmlResultOk) {
                    preparedMathML = buildPreparedMathML(inner, "MathJax.TeX.Parse");
                    pushAttempt({
                      source: "MathJax.TeX.Parse",
                      method: "toMathML",
                      ok: Boolean(preparedMathML),
                      latexLength: latexCandidate.length
                    });
                  }
                } else {
                  pushAttempt({ source: "MathJax.TeX.Parse", method: "toMathML", ok: false, error: "Parsed TeX did not expose toMathML()." });
                }
              } else {
                pushAttempt({ source: "MathJax.TeX.Parse", method: "availability", ok: false, error: "MathJax TeX parser unavailable." });
              }
            } catch (error) {
              pushAttempt({ source: "MathJax.TeX.Parse", method: "toMathML", ok: false, error: error && error.message ? error.message : String(error) });
            }
          }

          const runClassSync = (source) => invokeNoArgMethods(formulaEditorClass, source, [
            "updateByTextAreas",
            "redrawAll",
            "cleanupTextareas",
            "cleanupEditors"
          ]);

          const runEditorSync = (editor, source) => invokeNoArgMethods(editor, source, [
            "update",
            "redraw",
            "draw",
            "refresh",
            "render",
            "rebuild",
            "repaint",
            "save"
          ]);

          setTextareaValue(rawFormula, "textarea-seed");
          runClassSync("FormulaEditor.class-sync-before");

          const methodPriority = [
            "insertText",
            "insertString",
            "setText",
            "setValue",
            "setExpression",
            "setExpressionString",
            "setExpressionFromString",
            "setInput",
            "setMath",
            "loadText",
            "loadLatex",
            "fromLatex",
            "parse",
            "write"
          ];

          const tryMethod = (editorRef, source, methodName, args) => {
            if (!editorRef || typeof editorRef[methodName] !== "function") {
              return false;
            }
            try {
              editorRef[methodName](...args);
              pushAttempt({ source, method: methodName, ok: true, argsShape: args.map((arg) => (arg === null ? "null" : typeof arg)) });
              return true;
            } catch (error) {
              pushAttempt({ source, method: methodName, ok: false, error: error && error.message ? error.message : String(error) });
              return false;
            }
          };

          const keyCodeForChar = (ch) => {
            if (!ch) {
              return 0;
            }
            if (/^[0-9]$/.test(ch)) {
              return ch.charCodeAt(0);
            }
            if (/^[a-z]$/i.test(ch)) {
              return ch.charCodeAt(0);
            }
            const map = {
              " ": 32,
              "+": 43,
              "-": 45,
              "*": 42,
              "/": 47,
              "^": 94,
              "(": 40,
              ")": 41,
              ".": 46,
              ",": 44,
              "=": 61,
              "<": 60,
              ">": 62
            };
            return map[ch] || ch.charCodeAt(0) || 0;
          };

          const sendEditorKeydown = (editorRef, source, key, keyCode, meta) => {
            if (!editorRef || typeof editorRef.onkeydown !== "function") {
              return false;
            }
            try {
              editorRef.onkeydown({
                key,
                code: key,
                keyCode,
                which: keyCode,
                charCode: 0,
                preventDefault: () => {},
                stopPropagation: () => {}
              });
              pushAttempt({ source, method: `onkeydown(${key})`, ok: true, ...(meta || {}) });
              return true;
            } catch (error) {
              pushAttempt({ source, method: `onkeydown(${key})`, ok: false, error: error && error.message ? error.message : String(error), ...(meta || {}) });
              return false;
            }
          };

          const tryLoadMathMLIntoEditor = (editorRef, source) => {
            if (!preparedMathML || !editorRef || typeof editorRef.loadMathML !== "function") {
              return false;
            }

            const candidates = [];
            if (preparedMathML.string) {
              candidates.push({ label: "string", value: preparedMathML.string });
            }
            if (preparedMathML.node) {
              candidates.push({ label: "node", value: preparedMathML.node });
            }

            for (const candidate of candidates) {
              try {
                if (typeof editorRef.clearEditor === "function") {
                  editorRef.clearEditor();
                  pushAttempt({ source, method: "clearEditor(beforeLoadMathML)", ok: true });
                }
              } catch (error) {
                pushAttempt({ source, method: "clearEditor(beforeLoadMathML)", ok: false, error: error && error.message ? error.message : String(error) });
              }

              try {
                editorRef.loadMathML(candidate.value);
                pushAttempt({ source, method: "loadMathML", ok: true, inputType: candidate.label });
              } catch (error) {
                pushAttempt({ source, method: "loadMathML", ok: false, inputType: candidate.label, error: error && error.message ? error.message : String(error) });
                continue;
              }

              runEditorSync(editorRef, `${source}.sync-loadMathML`);
              runClassSync("FormulaEditor.class-sync-after-loadMathML");
              const changed = changedFrom(beforeState);
              if (changed.ok) {
                return true;
              }
            }

            return false;
          };

          const editorEntries = [...potentialEditors].filter((entry) => entry.editor && typeof entry.editor === "object");

          for (const entry of editorEntries) {
            if (tryLoadMathMLIntoEditor(entry.editor, entry.source)) {
              const changed = changedFrom(beforeState);
              return {
                ok: true,
                valueChanged: changed.valueChanged,
                canvasChanged: changed.canvasChanged,
                anyVisibleCanvasChanged: changed.anyVisibleCanvasChanged,
                beforeLength: beforeState.value.length,
                afterLength: changed.next.value.length,
                attemptedMethods: attempts,
                orgMathDoxAvailable: Boolean(orgFormulaEditor),
                slotTag: bundle.slot.tagName || null,
                slotClassName: bundle.slot.className || null,
                textareaClassName: bundle.textarea ? (bundle.textarea.className || null) : null,
                discoveredEditorMethods: editorEntries.map((candidate) => ({
                  source: candidate.source,
                  methodsSample: methodList(candidate.editor).slice(0, 80)
                }))
              };
            }

            const methods = methodList(entry.editor);
            const dynamic = methods.filter((name) => {
              const lower = String(name || "").toLowerCase();
              if (!lower) {
                return false;
              }
              if (lower.startsWith("setup")) {
                return false;
              }
              if (/(insert|load|from|parse|write|replace)/.test(lower)) {
                return true;
              }
              if (lower.startsWith("set") && /(text|value|expr|latex|mathml|input|string)/.test(lower)) {
                return true;
              }
              return false;
            });
            const ordered = [...methodPriority, ...dynamic.filter((name) => !methodPriority.includes(name))];

            for (const method of ordered) {
              const methodLower = String(method || "").toLowerCase();
              if (methodLower.startsWith("setup")) {
                continue;
              }
              if (methodLower.includes("mathml") && !/^\s*</.test(rawFormula)) {
                continue;
              }
              if (methodLower === "load" && !/^\s*</.test(rawFormula)) {
                continue;
              }

              const invoked = tryMethod(entry.editor, entry.source, method, [rawFormula]) ||
                tryMethod(entry.editor, entry.source, method, [rawFormula, false]) ||
                tryMethod(entry.editor, entry.source, method, [rawFormula, bundle.textarea]) ||
                tryMethod(entry.editor, entry.source, method, [bundle.textarea, rawFormula]);

              if (!invoked) {
                continue;
              }

              runEditorSync(entry.editor, `${entry.source}.sync`);
              runClassSync("FormulaEditor.class-sync-after-method");
              const changed = changedFrom(beforeState);
              if (changed.ok) {
                return {
                  ok: true,
                  valueChanged: changed.valueChanged,
                  canvasChanged: changed.canvasChanged,
                  anyVisibleCanvasChanged: changed.anyVisibleCanvasChanged,
                  beforeLength: beforeState.value.length,
                  afterLength: changed.next.value.length,
                  attemptedMethods: attempts,
                  orgMathDoxAvailable: Boolean(orgFormulaEditor),
                  slotTag: bundle.slot.tagName || null,
                  slotClassName: bundle.slot.className || null,
                  textareaClassName: bundle.textarea ? (bundle.textarea.className || null) : null,
                  discoveredEditorMethods: editorEntries.map((candidate) => ({
                    source: candidate.source,
                    methodsSample: methodList(candidate.editor).slice(0, 80)
                  }))
                };
              }
            }
          }

          const rawFormulaHasCommands = /\\[A-Za-z]+/.test(rawFormula || "");
          const allowKeypressFallback = !rawFormulaHasCommands;

          for (const entry of editorEntries) {
            const editor = entry.editor;
            if (!editor || typeof editor.onkeypress !== "function") {
              continue;
            }

            if (!allowKeypressFallback) {
              pushAttempt({
                source: entry.source,
                method: "onkeypress",
                ok: false,
                skipped: true,
                reason: "Disabled for raw LaTeX commands without MathML conversion."
              });
              continue;
            }

            try {
              if (typeof editor.focus === "function") {
                editor.focus();
                pushAttempt({ source: entry.source, method: "focus", ok: true });
              }
            } catch (error) {
              pushAttempt({ source: entry.source, method: "focus", ok: false, error: error && error.message ? error.message : String(error) });
            }

            try {
              if (typeof editor.clearEditor === "function") {
                editor.clearEditor();
                pushAttempt({ source: entry.source, method: "clearEditor", ok: true });
              }
            } catch (error) {
              pushAttempt({ source: entry.source, method: "clearEditor", ok: false, error: error && error.message ? error.message : String(error) });
            }

            let anyKeyOk = false;
            const chars = [...rawFormula];
            let scriptState = {
              active: false,
              grouped: false,
              depth: 0,
              sourceOperator: null
            };
            let pendingContainerExitAfterScript = false;

            for (let index = 0; index < chars.length; index += 1) {
              const ch = chars[index];
              const next = index + 1 < chars.length ? chars[index + 1] : "";

              if (pendingContainerExitAfterScript) {
                if (ch === ")" || ch === "}") {
                  sendEditorKeydown(editor, entry.source, "ArrowRight", 39, {
                    reason: "script-exit-container",
                    trigger: ch
                  });
                }
                pendingContainerExitAfterScript = false;
              }

              const code = keyCodeForChar(ch);
              try {
                const keyEvent = {
                  key: ch,
                  charCode: code,
                  keyCode: code,
                  which: code,
                  shiftKey: /^[A-Z]$/.test(ch),
                  preventDefault: () => {},
                  stopPropagation: () => {}
                };
                editor.onkeypress(keyEvent);
                pushAttempt({ source: entry.source, method: "onkeypress", ok: true, key: ch, keyCode: code });
                anyKeyOk = true;
              } catch (error) {
                pushAttempt({ source: entry.source, method: "onkeypress", ok: false, key: ch, keyCode: code, error: error && error.message ? error.message : String(error) });
                anyKeyOk = false;
                break;
              }

              if (scriptState.active) {
                if (!scriptState.grouped) {
                  if (ch === "(" || ch === "{") {
                    scriptState.grouped = true;
                    scriptState.depth = 1;
                  } else {
                    const nextIsScriptAtom = /^[\p{L}0-9]$/u.test(next);
                    if (!nextIsScriptAtom) {
                      sendEditorKeydown(editor, entry.source, "ArrowRight", 39, { reason: "script-exit", operator: scriptState.sourceOperator });
                      if (next === ")" || next === "}") {
                        pendingContainerExitAfterScript = true;
                      }
                      scriptState = { active: false, grouped: false, depth: 0, sourceOperator: null };
                    }
                  }
                } else {
                  if (ch === "(" || ch === "{") {
                    scriptState.depth += 1;
                  } else if (ch === ")" || ch === "}") {
                    scriptState.depth -= 1;
                  }

                  if (scriptState.depth <= 0) {
                    sendEditorKeydown(editor, entry.source, "ArrowRight", 39, { reason: "script-exit-group", operator: scriptState.sourceOperator });
                    if (next === ")" || next === "}") {
                      pendingContainerExitAfterScript = true;
                    }
                    scriptState = { active: false, grouped: false, depth: 0, sourceOperator: null };
                  }
                }
              } else if (ch === "_" || ch === "^") {
                scriptState = {
                  active: true,
                  grouped: false,
                  depth: 0,
                  sourceOperator: ch
                };
              }
            }

            if (!anyKeyOk) {
              continue;
            }

            if (scriptState.active) {
              sendEditorKeydown(editor, entry.source, "ArrowRight", 39, { reason: "script-exit-final", operator: scriptState.sourceOperator });
            }
            if (pendingContainerExitAfterScript) {
              sendEditorKeydown(editor, entry.source, "ArrowRight", 39, { reason: "script-exit-container-final" });
              pendingContainerExitAfterScript = false;
            }

            try {
              sendEditorKeydown(editor, entry.source, "Enter", 13);
            } catch (error) {
              pushAttempt({ source: entry.source, method: "onkeydown(Enter)", ok: false, error: error && error.message ? error.message : String(error) });
            }

            runEditorSync(editor, `${entry.source}.sync-keypress`);
            runClassSync("FormulaEditor.class-sync-after-keypress");

            const changed = changedFrom(beforeState);
            if (changed.ok) {
              return {
                ok: true,
                valueChanged: changed.valueChanged,
                canvasChanged: changed.canvasChanged,
                anyVisibleCanvasChanged: changed.anyVisibleCanvasChanged,
                beforeLength: beforeState.value.length,
                afterLength: changed.next.value.length,
                attemptedMethods: attempts,
                orgMathDoxAvailable: Boolean(orgFormulaEditor),
                slotTag: bundle.slot.tagName || null,
                slotClassName: bundle.slot.className || null,
                textareaClassName: bundle.textarea ? (bundle.textarea.className || null) : null,
                discoveredEditorMethods: editorEntries.map((candidate) => ({
                  source: candidate.source,
                  methodsSample: methodList(candidate.editor).slice(0, 80)
                }))
              };
            }
          }

          if (bundle.textarea) {
            bundle.textarea.focus();
            if (typeof bundle.textarea.setSelectionRange === "function") {
              bundle.textarea.setSelectionRange(0, bundle.textarea.value.length);
            }
            try {
              const execOk = document.execCommand("insertText", false, rawFormula);
              pushAttempt({ source: "textarea", method: "execCommand(insertText)", ok: execOk });
            } catch (error) {
              pushAttempt({ source: "textarea", method: "execCommand(insertText)", ok: false, error: error && error.message ? error.message : String(error) });
            }
          }

          runClassSync("FormulaEditor.class-sync-final");

          const changed = changedFrom(beforeState);
          const finalOk = hasLatexCommands ? false : changed.ok;
          const finalError = finalOk
            ? undefined
            : (hasLatexCommands
              ? "Complex LaTeX insertion failed: MathML import was not accepted by MathDox."
              : "No visible canvas change after MAIN world API attempts.");

          return {
            ok: finalOk,
            error: finalError,
            valueChanged: changed.valueChanged,
            canvasChanged: changed.canvasChanged,
            anyVisibleCanvasChanged: changed.anyVisibleCanvasChanged,
            beforeLength: beforeState.value.length,
            afterLength: changed.next.value.length,
            attemptedMethods: attempts,
            orgMathDoxAvailable: Boolean(orgFormulaEditor),
            slotTag: bundle.slot.tagName || null,
            slotClassName: bundle.slot.className || null,
            textareaClassName: bundle.textarea ? (bundle.textarea.className || null) : null,
            discoveredEditorMethods: potentialEditors.map((entry) => ({
              source: entry.source,
              methodsSample: methodList(entry.editor).slice(0, 80)
            }))
          };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const debugFrames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          result: entry.result || null
        }));

        const successful = (results || []).find((entry) => entry && entry.result && entry.result.ok);
        if (successful) {
          resolve({ ok: true, debugFrames, ...successful.result });
          return;
        }

        const firstError = debugFrames
          .map((entry) => entry.result && entry.result.error)
          .find(Boolean);

        resolve({ ok: false, error: firstError || "MAIN world MathDox API insertion failed.", debugFrames });
      }
    );
  });
}

function injectContentScript(tabId) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        files: ["content.js"]
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        resolve({
          ok: true,
          injectedFrames: Array.isArray(results) ? results.map((entry) => entry.frameId) : []
        });
      }
    );
  });
}

function sendMessageToFrame(tabId, frameId, message) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, message, { frameId }, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, frameId, error: chrome.runtime.lastError.message });
        return;
      }

      if (response && response.ok) {
        resolve({ ok: true, frameId, response });
        return;
      }

      resolve({
        ok: false,
        frameId,
        error: response && response.error ? response.error : "No response from frame.",
        response
      });
    });
  });
}

function sowisoKeyboardInsert(tabId, text) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        args: [text],
        func: async (formulaText) => {
          function isVisible(el) {
            if (!(el instanceof HTMLElement)) {
              return false;
            }

            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
          }

          function normalizeLabel(value) {
            return (value || "").replace(/\s+/g, " ").trim().toLowerCase();
          }

          function elementTokens(el) {
            const tokens = new Set();
            const maybeAdd = (value) => {
              const norm = normalizeLabel(value);
              if (norm) {
                tokens.add(norm);
              }
            };

            maybeAdd(el.textContent);
            maybeAdd(el.getAttribute("aria-label"));
            maybeAdd(el.getAttribute("title"));

            if (el instanceof HTMLElement) {
              maybeAdd(el.dataset.key);
              maybeAdd(el.dataset.value);
            }

            return tokens;
          }

          function clickLikeUser(el) {
            const events = ["pointerdown", "mousedown", "mouseup", "click"];
            for (const type of events) {
              el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            }
          }

          function wait(ms) {
            return new Promise((resolve) => window.setTimeout(resolve, ms));
          }

          function isLikelyKeyboardKey(el) {
            if (!(el instanceof HTMLElement) || !isVisible(el)) {
              return false;
            }

            const rect = el.getBoundingClientRect();
            if (rect.width < 14 || rect.height < 14 || rect.width > 130 || rect.height > 130) {
              return false;
            }

            const className = normalizeLabel(el.className || "");
            const role = normalizeLabel(el.getAttribute("role") || "");
            const tag = normalizeLabel(el.tagName || "");
            const hasKeyboardClass = className.includes("rs_skip") || className.includes("keyboard") || className.includes("key");
            const isClickishTag = tag === "td" || tag === "button";
            return hasKeyboardClass || role === "button" || isClickishTag;
          }

          function focusActiveAnswerSlot() {
            const slotSelectors = [
              "td[class*='pre_input'] + td canvas.mathdoxformula",
              "td[class*='pre_input'] + td [tabindex]",
              "td[class*='pre_input'] + td",
              "table.input_table td.pre_input_text + td canvas.mathdoxformula",
              "table.input_table td.pre_input_text + td [tabindex]",
              "table.input_table td.pre_input_text + td",
              "canvas.mathdoxformula",
              "td.rs_skip canvas.mathdoxformula"
            ];

            for (const selector of slotSelectors) {
              const candidates = [...document.querySelectorAll(selector)].filter((el) => isVisible(el));
              if (candidates.length === 0) {
                continue;
              }

              const target = candidates[0];
              clickLikeUser(target);
              if (target instanceof HTMLElement) {
                target.focus();
              }
              return {
                selector,
                tag: target.tagName,
                className: target.className || null
              };
            }

            return null;
          }

          function getKeyCells(root) {
            if (!root) {
              return [];
            }
            const selectors = "td.rs_skip, td[class*='rs_'], td, button, [role='button'], [data-key], [data-value]";
            return [...root.querySelectorAll(selectors)].filter((el) => isLikelyKeyboardKey(el));
          }

          function getGlobalKeyCells() {
            const selectors = "td.rs_skip, td[class*='rs_'], td, button, [role='button'], [data-key], [data-value]";
            return [...document.querySelectorAll(selectors)]
              .filter((el) => isLikelyKeyboardKey(el))
              .filter((el) => {
                const rect = el.getBoundingClientRect();
                return rect.top >= window.innerHeight * 0.42;
              });
          }

          function findKeyboardRoot() {
            const allTables = [...document.querySelectorAll("table")];
            const tables = allTables.filter((table) => isVisible(table));
            const scored = tables
              .map((table) => {
                const keys = getKeyCells(table);
                const rect = table.getBoundingClientRect();
                const text = normalizeLabel(table.textContent || "");
                const hintHits = ["standard", "function", "logic", "vector", "abc", "units", "std", "fun", "vec", "uni"]
                  .filter((hint) => text.includes(hint))
                  .length;
                const bottomBias = rect.top >= window.innerHeight * 0.45 ? 25 : 0;
                return {
                  table,
                  keyCount: keys.length,
                  rect,
                  area: rect.width * rect.height,
                  score: keys.length + hintHits * 20 + bottomBias
                };
              })
              .filter((entry) => entry.keyCount >= 10)
              .sort((a, b) => {
                if (b.score !== a.score) return b.score - a.score;
                if (b.keyCount !== a.keyCount) return b.keyCount - a.keyCount;
                if (b.rect.top !== a.rect.top) return b.rect.top - a.rect.top;
                return a.area - b.area;
              });

            if (scored[0]) {
              return {
                root: scored[0].table,
                diagnostics: {
                  allTables: allTables.length,
                  visibleTables: tables.length,
                  chosenKeyCount: scored[0].keyCount,
                  chosenScore: scored[0].score
                }
              };
            }

            const containers = [...document.querySelectorAll("div, section, aside")]
              .filter((el) => isVisible(el))
              .map((el) => {
                const rect = el.getBoundingClientRect();
                const keyCount = getKeyCells(el).length;
                const text = normalizeLabel(el.textContent || "");
                const hintHits = ["standard", "function", "logic", "vector", "abc", "units", "std", "fun", "vec", "uni"]
                  .filter((hint) => text.includes(hint))
                  .length;
                const bottomBias = rect.top >= window.innerHeight * 0.45 ? 20 : 0;
                return {
                  el,
                  rect,
                  keyCount,
                  score: keyCount + hintHits * 18 + bottomBias
                };
              })
              .filter((entry) => entry.keyCount >= 10)
              .sort((a, b) => {
                if (b.score !== a.score) return b.score - a.score;
                if (b.keyCount !== a.keyCount) return b.keyCount - a.keyCount;
                return b.rect.top - a.rect.top;
              });

            if (containers[0]) {
              return {
                root: containers[0].el,
                diagnostics: {
                  allTables: allTables.length,
                  visibleTables: tables.length,
                  chosenKeyCount: containers[0].keyCount,
                  chosenScore: containers[0].score,
                  containerFallback: true
                }
              };
            }

            return {
              root: null,
              diagnostics: {
                allTables: allTables.length,
                visibleTables: tables.length
              }
            };
          }

          function findKeyboardShell(root) {
            const hints = ["std", "standard", "fun", "function", "log", "logic", "vec", "vector", "abc", "uni", "units"];
            let current = root;
            let best = root;
            let bestHits = 0;

            for (let i = 0; i < 6 && current; i += 1) {
              const text = normalizeLabel(current.textContent || "");
              const hits = hints.filter((hint) => text.includes(hint)).length;
              if (hits > bestHits) {
                bestHits = hits;
                best = current;
              }
              current = current.parentElement;
            }

            return best;
          }

          function findByTokens(candidates, tokenOptions) {
            const desired = tokenOptions.map((token) => normalizeLabel(token)).filter(Boolean);
            if (desired.length === 0) {
              return null;
            }

            const matches = candidates
              .map((el) => {
                const rect = el.getBoundingClientRect();
                return { el, rect, tokens: elementTokens(el) };
              })
              .filter((entry) => desired.some((token) => entry.tokens.has(token)))
              .sort((a, b) => {
                if (b.rect.top !== a.rect.top) return b.rect.top - a.rect.top;
                return a.rect.left - b.rect.left;
              });

            return matches[0] ? matches[0].el : null;
          }

          function findTabControl(shell, root, aliases) {
            const rootRect = root.getBoundingClientRect();
            const rootCandidates = [...root.querySelectorAll("td, a, button, [role='tab'], [role='button'], li, span, div")];
            const shellCandidates = [...shell.querySelectorAll("td, a, button, [role='tab'], [role='button'], li, span, div")];
            const candidates = [...new Set([...rootCandidates, ...shellCandidates])]
              .filter((el) => isVisible(el))
              .filter((el) => {
                const rect = el.getBoundingClientRect();
                const nearKeyboardTop = rect.bottom <= rootRect.top + 90 && rect.top >= rootRect.top - 180;
                const smallControl = rect.width >= 14 && rect.width <= 90 && rect.height >= 10 && rect.height <= 45;
                return nearKeyboardTop && smallControl;
              });

            return findByTokens(candidates, aliases);
          }

          function tokenOptionsForChar(ch) {
            if (ch === "*") return ["*", "×", "·", "⋅"];
            if (ch === "-") return ["-", "−"];
            if (ch === "/") return ["/", "÷", "frac", "fraction", "a/b"];
            if (ch === "^") return ["^", "pow", "power", "x^y", "sup", "superscript"];
            if (ch === "_") return ["_", "sub", "subscript", "x_y"];
            if (ch === "<") return ["<", "≤", "le"];
            if (ch === ">") return [">", "≥", "ge"];
            if (ch === "(") return ["("];
            if (ch === ")") return [")"];
            return [ch];
          }

          function preferredTabsForChar(ch) {
            if (/[a-z]/i.test(ch)) {
              return ["abc", "standard", "function"];
            }
            if (/\d/.test(ch)) {
              return ["standard", "abc"];
            }
            if (ch === "^" || ch === "_" || ch === "/") {
              return ["function", "abc", "standard"];
            }
            return ["standard", "function", "abc"];
          }

          function findLegacyCandidate(tokenOptions) {
            const desired = tokenOptions.map((token) => normalizeLabel(token)).filter(Boolean);
            if (desired.length === 0) {
              return null;
            }

            const selectors =
              "td.rs_skip, button, [role='button'], [role='tab'], [data-key], [data-value], [aria-label], .key, .keyboard-key, td, th, li, span";

            const candidates = [...document.querySelectorAll(selectors)]
              .filter((el) => isVisible(el))
              .filter((el) => {
                const rect = el.getBoundingClientRect();
                return rect.top >= window.innerHeight * 0.42 && rect.width >= 14 && rect.width <= 130 && rect.height >= 12 && rect.height <= 130;
              })
              .map((el) => ({ el, rect: el.getBoundingClientRect(), tokens: elementTokens(el) }))
              .filter((entry) => desired.some((token) => entry.tokens.has(token)))
              .sort((a, b) => {
                if (b.rect.top !== a.rect.top) return b.rect.top - a.rect.top;
                return a.rect.left - b.rect.left;
              });

            return candidates[0] ? candidates[0].el : null;
          }

          function codeForChar(ch) {
            if (/^[a-z]$/i.test(ch)) {
              return `Key${ch.toUpperCase()}`;
            }
            if (/^[0-9]$/.test(ch)) {
              return `Digit${ch}`;
            }
            const map = {
              "=": "Equal",
              "+": "Equal",
              "-": "Minus",
              "_": "Minus",
              "/": "Slash",
              "?": "Slash",
              ".": "Period",
              ",": "Comma",
              "(": "Digit9",
              ")": "Digit0",
              "^": "Digit6",
              "*": "Digit8",
              "<": "Comma",
              ">": "Period",
              "[": "BracketLeft",
              "]": "BracketRight"
            };
            return map[ch] || "Unidentified";
          }

          function fireKeyboardSequence(target, ch) {
            const code = codeForChar(ch);
            const upper = /^[A-Z]$/.test(ch);
            const key = ch;
            const send = (receiver, type, params) => {
              receiver.dispatchEvent(
                new KeyboardEvent(type, {
                  bubbles: true,
                  cancelable: true,
                  composed: true,
                  ...params
                })
              );
            };

            if (upper) {
              send(target, "keydown", { key: "Shift", code: "ShiftLeft", shiftKey: true });
            }

            send(target, "keydown", { key, code, shiftKey: upper });
            send(target, "keypress", { key, code, shiftKey: upper });

            try {
              target.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: ch, inputType: "insertText" }));
            } catch (_error) {
              // Ignore.
            }

            try {
              target.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: false, data: ch, inputType: "insertText" }));
            } catch (_error) {
              target.dispatchEvent(new Event("input", { bubbles: true }));
            }

            send(target, "keyup", { key, code, shiftKey: upper });

            if (upper) {
              send(target, "keyup", { key: "Shift", code: "ShiftLeft", shiftKey: false });
            }
          }

          function snapshotEditorState() {
            const pieces = [];
            const mathTextareas = [...document.querySelectorAll("textarea.math-editor, textarea.mathdoxformula, textarea[class*='mathdox'], textarea[class*='math-editor']")];
            for (const ta of mathTextareas) {
              pieces.push(`m:${ta.value || ""}`);
            }

            const visibleEditors = [
              ...document.querySelectorAll(
                "input, textarea, [contenteditable]:not([contenteditable='false']), [role='textbox'], .mq-editable-field, .mathquill-editable"
              )
            ].filter((el) => isVisible(el));

            for (const el of visibleEditors) {
              if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
                pieces.push(`v:${el.value || ""}`);
              } else {
                pieces.push(`v:${el.textContent || ""}`);
              }
            }

            return pieces.join("||");
          }

          const latexConversions = [
            [/\\cdot/g, "*"],
            [/\\times/g, "*"],
            [/\\left/g, ""],
            [/\\right/g, ""]
          ];

          let normalizedFormula = formulaText || "";
          for (const [pattern, replacement] of latexConversions) {
            normalizedFormula = normalizedFormula.replace(pattern, replacement);
          }

          const chars = [...normalizedFormula];
          const slotFocus = focusActiveAnswerSlot();
          const keyboardLookup = findKeyboardRoot();
          const keyboardRoot = keyboardLookup.root || document.body;
          const rootWasMissing = !keyboardLookup.root;

          const initialGlobalKeys = getGlobalKeyCells();
          const initialLocalKeys = getKeyCells(keyboardRoot);

          if (initialGlobalKeys.length < 10 && initialLocalKeys.length < 10) {
            return {
              ok: false,
              error: "Could not find the Sowiso keyboard container in this frame.",
              keyboardDiagnostics: {
                slotFocus,
                ...keyboardLookup.diagnostics,
                rootWasMissing,
                initialGlobalKeys: initialGlobalKeys.length,
                initialLocalKeys: initialLocalKeys.length
              }
            };
          }

          const keyboardShell = findKeyboardShell(keyboardRoot);
          function currentKeyCells() {
            const local = getKeyCells(keyboardRoot);
            const global = getGlobalKeyCells();
            if (local.length >= 20) {
              return local;
            }
            if (global.length > local.length) {
              return global;
            }
            return local;
          }

          const keyCells = currentKeyCells();
          if (keyCells.length < 10) {
            return {
              ok: false,
              error: "Could not resolve visible keyboard keys in this frame.",
              keyboardDiagnostics: {
                slotFocus,
                ...keyboardLookup.diagnostics,
                rootWasMissing,
                globalKeys: getGlobalKeyCells().length,
                localKeys: getKeyCells(keyboardRoot).length
              }
            };
          }

          const tabAliases = {
            standard: ["standard", "std"],
            abc: ["abc", "variables", "letters"],
            function: ["function", "fun", "logic", "log"],
            vector: ["vector", "vec"],
            units: ["units", "uni"]
          };

          const toggleShiftAliases = ["shift", "caps", "uppercase", "upper"];
          let activeTab = null;

          async function switchTab(tabName) {
            const aliases = tabAliases[tabName] || [tabName];
            let tabControl = findTabControl(keyboardShell, keyboardRoot, aliases);
            if (!tabControl) {
              const aliasSet = new Set(aliases.map((item) => normalizeLabel(item)));
              const globalCandidates = [...document.querySelectorAll("td, a, button, [role='tab'], [role='button'], li, span, div")]
                .filter((el) => isVisible(el))
                .filter((el) => {
                  const rect = el.getBoundingClientRect();
                  if (rect.top < window.innerHeight * 0.45) {
                    return false;
                  }
                  if (rect.width < 14 || rect.width > 90 || rect.height < 10 || rect.height > 45) {
                    return false;
                  }
                  const tokens = elementTokens(el);
                  for (const token of tokens) {
                    if (aliasSet.has(token)) {
                      return true;
                    }
                  }
                  return false;
                })
                .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);

              tabControl = globalCandidates[0] || null;
            }

            if (!tabControl) {
              return false;
            }
            clickLikeUser(tabControl);
            activeTab = tabName;
            await wait(70);
            return true;
          }

          function findKey(tokenOptions) {
            return findByTokens(currentKeyCells(), tokenOptions);
          }

          async function toggleShiftIfNeeded() {
            const shiftKey = findKey(toggleShiftAliases);
            if (!shiftKey) {
              return false;
            }
            clickLikeUser(shiftKey);
            await wait(30);
            return true;
          }

          const baseline = snapshotEditorState();
          let insertedCount = 0;
          const failed = [];
          const trace = [];

          for (const ch of chars) {
            if (ch === " " || ch === "\n" || ch === "\t") {
              continue;
            }

            const before = snapshotEditorState();
            const tabsToTry = preferredTabsForChar(ch);
            const needsShift = /[A-Z]/.test(ch);
            let button = null;
            let usedTab = null;
            let usedShift = false;
            const tabSwitchReport = [];

            for (const tab of tabsToTry) {
              if (activeTab !== tab) {
                const switched = await switchTab(tab);
                tabSwitchReport.push({ tab, switched });
              } else {
                tabSwitchReport.push({ tab, switched: true, alreadyActive: true });
              }

              if (needsShift && tab === "abc") {
                usedShift = await toggleShiftIfNeeded();
                button = findKey([ch, ch.toLowerCase()]);
              } else {
                button = findKey(tokenOptionsForChar(ch));
                if (!button && /[a-z]/i.test(ch)) {
                  button = findKey([ch.toLowerCase(), ch.toUpperCase()]);
                }
              }

              if (button) {
                usedTab = tab;
                break;
              }
            }

            if (!button) {
              const legacyOptions = /[a-z]/i.test(ch)
                ? [...tokenOptionsForChar(ch), ch.toLowerCase(), ch.toUpperCase()]
                : tokenOptionsForChar(ch);
              button = findLegacyCandidate(legacyOptions);
              if (button) {
                usedTab = "legacy";
              }
            }

            if (!button) {
              const activeTarget = document.activeElement instanceof HTMLElement ? document.activeElement : document.body;
              fireKeyboardSequence(activeTarget, ch);
              fireKeyboardSequence(document, ch);
              await wait(25);

              const typedAfter = snapshotEditorState();
              const typedChanged = typedAfter !== before;
              if (typedChanged) {
                insertedCount += 1;
                trace.push({
                  char: ch,
                  ok: true,
                  method: "typed-fallback",
                  tabsTried: tabsToTry,
                  tabSwitchReport,
                  activeTag: activeTarget ? activeTarget.tagName : null
                });
                continue;
              }

              failed.push(ch);
              trace.push({ char: ch, ok: false, reason: "no-matching-key", tabsTried: tabsToTry, tabSwitchReport });
              continue;
            }

            clickLikeUser(button);
            insertedCount += 1;
            await wait(35);

            const after = snapshotEditorState();
            const changed = before !== after;
            trace.push({
              char: ch,
              ok: changed,
              tab: usedTab,
              shift: usedShift,
              keyText: normalizeLabel(button.textContent || ""),
              keyAria: normalizeLabel(button.getAttribute("aria-label") || ""),
              keyClass: button.className || null
            });
          }

          await wait(90);
          const changed = snapshotEditorState() !== baseline;

          if (insertedCount > 0 && changed) {
            return {
              ok: true,
              insertedCount,
              failed: failed.join(""),
              changed,
              trace,
              keyboardDiagnostics: {
                slotFocus,
                ...keyboardLookup.diagnostics,
                rootWasMissing,
                globalKeys: getGlobalKeyCells().length,
                localKeys: getKeyCells(keyboardRoot).length
              }
            };
          }

          if (insertedCount > 0 && !changed) {
            return {
              ok: false,
              error: "Keyboard keys were clicked, but no editor value changed. Click the pink input slot first.",
              insertedCount,
              failed: failed.join(""),
              changed,
              trace,
              keyboardDiagnostics: {
                slotFocus,
                ...keyboardLookup.diagnostics,
                rootWasMissing,
                globalKeys: getGlobalKeyCells().length,
                localKeys: getKeyCells(keyboardRoot).length
              }
            };
          }

          return {
            ok: false,
              error: failed.length > 0 ? `Unsupported characters for keyboard input: ${failed.join("")}` : "No keys were inserted.",
            trace,
            keyboardDiagnostics: {
              slotFocus,
              ...keyboardLookup.diagnostics,
              rootWasMissing,
              globalKeys: getGlobalKeyCells().length,
              localKeys: getKeyCells(keyboardRoot).length
            }
          };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const debugFrames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          result: entry.result || null
        }));

        const successful = (results || []).find((entry) => entry && entry.result && entry.result.ok);
        if (successful) {
          resolve({ ...successful.result, debugFrames });
          return;
        }

        const firstError = debugFrames
          .map((entry) => entry.result && entry.result.error)
          .find(Boolean);

        resolve({ ok: false, error: firstError || "Sowiso keyboard insertion failed.", debugFrames });
      }
    );
  });
}

function directInsertFallback(tabId, formula) {
  return new Promise((resolve) => {
    if (!chrome.scripting || typeof chrome.scripting.executeScript !== "function") {
      resolve({ ok: false, error: "Scripting API unavailable." });
      return;
    }

    chrome.scripting.executeScript(
      {
        target: { tabId, allFrames: true },
        args: [formula],
        func: (text) => {
          const selectors = [
            "textarea",
            "input:not([type])",
            "input[type='text']",
            "input[type='search']",
            "input[type='email']",
            "input[type='url']",
            "input[type='tel']",
            "input[type='number']",
            "input[type='password']",
            "[contenteditable]:not([contenteditable='false'])",
            "[role='textbox']",
            ".mq-editable-field",
            ".mathquill-editable"
          ];

          function isVisible(el) {
            if (!(el instanceof HTMLElement)) {
              return false;
            }

            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
          }

          function isTextInput(el) {
            if (!(el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement)) {
              return false;
            }

            if (el instanceof HTMLTextAreaElement) {
              return true;
            }

            const t = (el.type || "text").toLowerCase();
            return ["text", "search", "email", "url", "tel", "number", "password"].includes(t);
          }

          function isLikelySowisoMathTextarea(el) {
            if (!(el instanceof HTMLTextAreaElement)) {
              return false;
            }

            const className = (el.className || "").toLowerCase();
            return className.includes("math-editor") || className.includes("mathdoxformula");
          }

          function readValue(el) {
            if (isTextInput(el)) {
              return el.value || "";
            }

            if (el && el.isContentEditable) {
              return el.textContent || "";
            }

            return "";
          }

          function normalize(el) {
            if (!el || !(el instanceof Element)) {
              return null;
            }

            if (isTextInput(el)) {
              return el;
            }

            if (el.isContentEditable) {
              return el;
            }

            return (
              el.closest("[contenteditable]:not([contenteditable='false'])") ||
              el.closest("[role='textbox']") ||
              null
            );
          }

          function insertTextInput(el, value) {
            el.focus();
            const start = el.selectionStart ?? el.value.length;
            const end = el.selectionEnd ?? start;
            if (typeof el.setRangeText === "function") {
              el.setRangeText(value, start, end, "end");
            } else {
              el.value = `${el.value.slice(0, start)}${value}${el.value.slice(end)}`;
            }
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
          }

          function insertContentEditable(el, value) {
            el.focus();
            try {
              if (document.execCommand("insertText", false, value)) {
                el.dispatchEvent(new Event("input", { bubbles: true }));
                el.dispatchEvent(new Event("change", { bubbles: true }));
                return;
              }
            } catch (_error) {}

            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) {
              el.append(document.createTextNode(value));
            } else {
              const range = sel.getRangeAt(0);
              range.deleteContents();
              const node = document.createTextNode(value);
              range.insertNode(node);
              range.setStartAfter(node);
              range.setEndAfter(node);
              sel.removeAllRanges();
              sel.addRange(range);
            }

            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
          }

          const activeCandidate = normalize(document.activeElement);
          const visibleCandidates = [...document.querySelectorAll(selectors.join(","))]
            .map((el) => normalize(el))
            .filter((el) => el && isVisible(el));

          const candidate = activeCandidate && isVisible(activeCandidate) ? activeCandidate : visibleCandidates[0] || null;

          if (!candidate) {
            return { ok: false, error: "No editor target in this frame." };
          }

          if (isLikelySowisoMathTextarea(candidate)) {
            return { ok: false, error: "Direct fallback reached hidden MathDox textarea." };
          }

          const before = readValue(candidate);

          if (isTextInput(candidate)) {
            insertTextInput(candidate, text);
            const after = readValue(candidate);
            if (after !== before) {
              return { ok: true, method: "text", beforeLength: before.length, afterLength: after.length };
            }
            return { ok: false, error: "Direct text insert did not change the target value." };
          }

          if (candidate.isContentEditable) {
            insertContentEditable(candidate, text);
            const after = readValue(candidate);
            if (after !== before) {
              return { ok: true, method: "contenteditable", beforeLength: before.length, afterLength: after.length };
            }
            return { ok: false, error: "Direct contenteditable insert did not change the target value." };
          }

          return { ok: false, error: "Unsupported target in this frame." };
        }
      },
      (results) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        const debugFrames = (results || []).map((entry) => ({
          frameId: entry.frameId,
          result: entry.result || null
        }));

        const successful = (results || []).find((entry) => entry && entry.result && entry.result.ok);
        if (successful) {
          resolve({ ok: true, debugFrames });
          return;
        }

        const firstError = debugFrames
          .map((entry) => entry.result && entry.result.error)
          .find(Boolean);

        resolve({ ok: false, error: firstError || "Direct insertion fallback failed.", debugFrames });
      }
    );
  });
}

async function sendInsertToAnyFrame(tabId, formula, preferredFrameIds) {
  const frameDetails = await getFrameDetails(tabId);
  const preferred = Array.isArray(preferredFrameIds) && preferredFrameIds.length > 0
    ? preferredFrameIds
    : frameDetails.map((frame) => frame.frameId);
  const frameIds = [...new Set(preferred)];
  const fatalErrors = [];
  const ignoredErrors = [];
  const attempts = [];

  for (const frameId of frameIds) {
    const result = await sendMessageToFrame(tabId, frameId, { type: "INSERT_LATEX", formula });
    attempts.push(result);
    if (result.ok) {
      return { ok: true, attempts, frameIds, frameDetails };
    }

    if (result.error) {
      if (result.error.includes("Receiving end does not exist")) {
        ignoredErrors.push(result.error);
      } else {
        fatalErrors.push(result.error);
      }
    }
  }

  const fallbackMessage =
    fatalErrors[fatalErrors.length - 1] ||
    ignoredErrors[ignoredErrors.length - 1] ||
    "Insertion failed in all frames.";
  return { ok: false, error: fallbackMessage, attempts, frameIds, frameDetails };
}

async function insertIntoPage() {
  const raw = latexInput.value.trim();
  appendDebug("Insert requested", { raw, mode: insertMode.value, sowisoConvert: sowisoConvert.checked });

  if (!raw) {
    setStatus("Type a LaTeX formula first.", true);
    appendDebug("Aborted: empty formula");
    return;
  }

  const wrapped = wrapFormula(raw, insertMode.value);
  const autoConvert = false;
  const effectiveConvert = sowisoConvert.checked;
  const formula = effectiveConvert ? convertLatexToSowisoLinear(wrapped) : wrapped;
  appendDebug("Prepared formula", { wrapped, formula, effectiveConvert, autoConvert });
  if (autoConvert) {
    appendDebug("Auto-converted complex LaTeX for Sowiso input");
  }
  const activeTab = await getActiveTab();
  appendDebug("Active tab", {
    id: activeTab && activeTab.id,
    url: activeTab && activeTab.url
  });

  if (!activeTab || activeTab.id === undefined) {
    setStatus("No active browser tab found.", true);
    appendDebug("Aborted: no active tab");
    return;
  }

  if (!isInjectableTabUrl(activeTab.url)) {
    setStatus("Open a regular website tab first.", true);
    appendDebug("Aborted: unsupported tab URL", { url: activeTab.url || null });
    return;
  }

  const isSowisoTab = isSowisoTabUrl(activeTab.url);
  const frameDetails = await getFrameDetails(activeTab.id);
  appendDebug("Frame details", frameDetails);

  const injection = await injectContentScript(activeTab.id);
  appendDebug("Content script injection", injection);

  let slotFocus = { ok: false, skipped: !isSowisoTab };
  let mathdoxResult = { ok: false, skipped: !isSowisoTab };
  let slotKeyTypeResult = { ok: false, skipped: !isSowisoTab };
  let mainWorldApiResult = { ok: false, skipped: !isSowisoTab };

  if (isSowisoTab) {
    slotFocus = await focusSowisoSlot(activeTab.id);
    appendDebug("Slot focus attempt", slotFocus);

    mathdoxResult = await sowisoTextareaInsert(activeTab.id, formula);
    appendDebug("MathDox textarea insert result", mathdoxResult);
    if (mathdoxResult.ok) {
      setStatus("Formula inserted.", false);
      appendDebug("Completed via MathDox textarea insertion");
      return;
    }

    slotKeyTypeResult = await sowisoSlotKeyTypeInsert(activeTab.id, formula);
    appendDebug("Direct slot key typing result", slotKeyTypeResult);
    if (slotKeyTypeResult.ok) {
      setStatus("Formula inserted.", false);
      appendDebug("Completed via direct slot key typing");
      return;
    }

    mainWorldApiResult = await sowisoMainWorldApiInsert(activeTab.id, formula, wrapped);
    appendDebug("MAIN world MathDox API result", mainWorldApiResult);
    if (mainWorldApiResult.ok) {
      setStatus("Formula inserted.", false);
      appendDebug("Completed via MAIN world MathDox API");
      return;
    }
  } else {
    appendDebug("Skipping Sowiso-specific insertion flow", { url: activeTab.url || null });
  }

  const frameInsertResult = await sendInsertToAnyFrame(activeTab.id, formula, injection.injectedFrames);
  appendDebug("Content script direct insert result", frameInsertResult);
  if (frameInsertResult.ok) {
    setStatus("Formula inserted.", false);
    appendDebug("Completed via content script insertion");
    return;
  }

  const directFallbackResult = await directInsertFallback(activeTab.id, formula);
  appendDebug("Direct insertion fallback result", directFallbackResult);
  if (directFallbackResult.ok) {
    setStatus("Formula inserted.", false);
    appendDebug("Completed via direct insertion fallback");
    return;
  }

  const failureMessage =
    frameInsertResult.error ||
    directFallbackResult.error ||
    mathdoxResult.error ||
    slotKeyTypeResult.error ||
    mainWorldApiResult.error ||
    "Direct insertion failed.";
  setStatus(failureMessage, true);
  appendDebug("Stopped: all direct insertion paths failed", {
    mathdoxResult,
    slotKeyTypeResult,
    mainWorldApiResult,
    frameInsertResult,
    directFallbackResult
  });
}

previewImage.addEventListener("load", () => {
  previewButton.classList.add("has-image");
  setStatus("", false);
});

previewImage.addEventListener("error", () => {
  showEmptyPreview();
  previewPlaceholder.textContent = "Preview failed. Check your LaTeX syntax.";
  setStatus("Could not render preview image.", true);
});

for (const chip of themeChips) {
  chip.addEventListener("click", () => {
    const nextMode = chip.dataset.theme || "auto";
    applyTheme(nextMode);
    persistSettings({ themeMode: nextMode });
    updatePreview();
  });
}

if (sowisoConvert) {
  sowisoConvert.addEventListener("change", () => {
    appendDebug("Updated setting", { sowisoConvert: sowisoConvert.checked });
  });
}

latexInput.addEventListener("input", updatePreview);
insertButton.addEventListener("click", () => {
  insertIntoPage();
});
previewButton.addEventListener("click", () => {
  insertIntoPage();
});

clearButton.addEventListener("click", () => {
  latexInput.value = "";
  showEmptyPreview();
  setStatus("", false);
});

if (clearDebugButton) {
  clearDebugButton.addEventListener("click", () => {
    debugLines = [];
    if (debugLogEl) {
      debugLogEl.textContent = "";
    }
    appendDebug("Debug log cleared");
  });
}

if (copyDebugButton) {
  copyDebugButton.addEventListener("click", () => {
    copyDebugLog();
  });
}

latexInput.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    event.preventDefault();
    insertIntoPage();
  }
});

systemTheme.addEventListener("change", () => {
  if (currentThemeMode === "auto") {
    applyTheme("auto");
    updatePreview();
  }
});

chrome.storage.local.get(DEFAULT_SETTINGS, (settings) => {
  const initialMode = settings.themeMode || "auto";
  applyTheme(initialMode);
  if (sowisoConvert) {
    sowisoConvert.checked = false;
  }
  updatePreview();
  appendDebug("Panel ready", { theme: initialMode, sowisoConvert: false });
});

showEmptyPreview();
