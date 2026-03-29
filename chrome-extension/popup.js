const latexInput = document.getElementById("latexInput");
const insertMode = document.getElementById("insertMode");
const sowisoConvert = document.getElementById("sowisoConvert");
const previewButton = document.getElementById("previewButton");
const previewImage = document.getElementById("previewImage");
const previewPlaceholder = document.getElementById("previewPlaceholder");
const insertButton = document.getElementById("insertButton");
const clearButton = document.getElementById("clearButton");
const statusEl = document.getElementById("status");
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
    const replacement = `((${num})/(${den}))`;
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

function convertLatexToSowisoLinear(input) {
  const greekMap = {
    "\\alpha": "alpha",
    "\\beta": "beta",
    "\\gamma": "gamma",
    "\\delta": "delta",
    "\\epsilon": "epsilon",
    "\\theta": "theta",
    "\\lambda": "lambda",
    "\\mu": "mu",
    "\\pi": "pi",
    "\\rho": "rho",
    "\\sigma": "sigma",
    "\\tau": "tau",
    "\\phi": "phi",
    "\\omega": "omega"
  };

  let out = input || "";
  out = out.replace(/^\s*\$+|\$+\s*$/g, "");
  out = out.replace(/\\left|\\right|\\,/g, "");
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

  // Handle common grouped powers/subscripts.
  out = out.replace(/\^\{([^{}]+)\}/g, "^($1)");
  out = out.replace(/_\{([^{}]+)\}/g, "_($1)");

  out = out.replace(/[{}]/g, (ch) => (ch === "{" ? "(" : ")"));
  out = out.replace(/\\([a-zA-Z]+)/g, "$1");
  out = out.replace(/\s+/g, "");
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

          const selectors = [
            "td[class*='pre_input'] + td canvas.mathdoxformula",
            "td[class*='pre_input'] + td [tabindex]",
            "td[class*='pre_input'] + td",
            "table.input_table td.pre_input_text + td canvas.mathdoxformula",
            "table.input_table td.pre_input_text + td [tabindex]",
            "table.input_table td.pre_input_text + td",
            "canvas.mathdoxformula"
          ];

          for (const selector of selectors) {
            const nodes = [...document.querySelectorAll(selector)].filter((el) => isVisible(el));
            if (nodes.length === 0) {
              continue;
            }

            const target = nodes[0];
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

          function findTextareaForActiveSlot() {
            const active = document.activeElement;
            if (active instanceof HTMLElement && active.classList.contains("mathdoxformula")) {
              const td = active.closest("td");
              if (td) {
                const ta = td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']");
                if (ta instanceof HTMLTextAreaElement) {
                  return ta;
                }
              }
            }

            const slot = [...document.querySelectorAll("td[class*='pre_input'] + td canvas.mathdoxformula, table.input_table td.pre_input_text + td canvas.mathdoxformula, canvas.mathdoxformula")]
              .find((el) => isVisible(el));
            if (slot) {
              clickLikeUser(slot);
              if (slot instanceof HTMLElement) {
                slot.focus();
              }
              const td = slot.closest("td");
              if (td) {
                const ta = td.querySelector("textarea.mathdoxformula, textarea.math-editor, textarea[class*='mathdox'], textarea[class*='math-editor']");
                if (ta instanceof HTMLTextAreaElement) {
                  return ta;
                }
              }
            }

            return null;
          }

          const target = findTextareaForActiveSlot();
          if (!target) {
            return { ok: false, error: "Could not find MathDox textarea for active answer slot." };
          }

          const before = target.value || "";
          target.focus();

          if (typeof target.setRangeText === "function") {
            target.setRangeText(rawFormula, 0, target.value.length, "end");
          } else {
            target.value = rawFormula;
          }

          try {
            target.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: rawFormula, inputType: "insertText" }));
          } catch (_error) {
            // Ignore.
          }

          target.dispatchEvent(new Event("input", { bubbles: true }));
          target.dispatchEvent(new Event("change", { bubbles: true }));
          target.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, key: "Unidentified", code: "Unidentified" }));
          target.dispatchEvent(new Event("blur", { bubbles: true }));

          if (window.jQuery && typeof window.jQuery === "function") {
            try {
              window.jQuery(target).trigger("input");
              window.jQuery(target).trigger("change");
              window.jQuery(target).trigger("keyup");
            } catch (_error) {
              // Ignore.
            }
          }

          return {
            ok: target.value !== before || rawFormula.length === 0,
            beforeLength: before.length,
            afterLength: (target.value || "").length,
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

async function sendInsertToAnyFrame(tabId, formula) {
  const frameDetails = await getFrameDetails(tabId);
  const frameIds = [...new Set(frameDetails.map((frame) => frame.frameId))];
  const lastErrors = [];
  const attempts = [];

  for (const frameId of frameIds) {
    const result = await sendMessageToFrame(tabId, frameId, { type: "INSERT_LATEX", formula });
    attempts.push(result);
    if (result.ok) {
      return { ok: true, attempts, frameIds, frameDetails };
    }

    if (result.error) {
      lastErrors.push(result.error);
    }
  }

  const fallbackMessage = lastErrors[lastErrors.length - 1] || "Insertion failed in all frames.";
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
  const formula = sowisoConvert.checked ? convertLatexToSowisoLinear(wrapped) : wrapped;
  appendDebug("Prepared formula", { wrapped, formula });
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

  if (!activeTab.url || !activeTab.url.startsWith("https://cloud.sowiso.nl/")) {
    setStatus("Open a cloud.sowiso.nl exercise tab first.", true);
    appendDebug("Aborted: unsupported tab URL", { url: activeTab.url || null });
    return;
  }

  const frameDetails = await getFrameDetails(activeTab.id);
  appendDebug("Frame details", frameDetails);

  const injection = await injectContentScript(activeTab.id);
  appendDebug("Content script injection", injection);

  const slotFocus = await focusSowisoSlot(activeTab.id);
  appendDebug("Slot focus attempt", slotFocus);

  const mathdoxResult = await sowisoTextareaInsert(activeTab.id, formula);
  appendDebug("MathDox textarea insert result", mathdoxResult);
  if (mathdoxResult.ok) {
    setStatus("Formula inserted.", false);
    appendDebug("Completed via MathDox textarea insertion");
    return;
  }

  const debuggerResult = await debuggerTypeInsert(activeTab.id, formula);
  appendDebug("Debugger typing result", debuggerResult);
  if (debuggerResult.ok) {
    setStatus("Formula inserted.", false);
    appendDebug("Completed via debugger typing");
    return;
  }

  const keyboardResult = await sowisoKeyboardInsert(activeTab.id, formula);
  appendDebug("Keyboard insert result", keyboardResult);
  if (keyboardResult.ok) {
    if (keyboardResult.failed) {
      setStatus(`Partially inserted. Missing: ${keyboardResult.failed}`, true);
      appendDebug("Keyboard insert partial", { failed: keyboardResult.failed });
      return;
    }

    setStatus("Formula inserted.", false);
    appendDebug("Completed via keyboard simulation");
    return;
  }

  setStatus(keyboardResult.error || "Insertion failed.", true);
  appendDebug("Stopped: keyboard path failed (fallbacks disabled for safety)", keyboardResult);
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
