(function initSowisoHelperContent() {
if (globalThis.__SOWISO_HELPER_CONTENT_V1__) {
  return;
}

globalThis.__SOWISO_HELPER_CONTENT_V1__ = true;

const ALLOWED_INPUT_TYPES = new Set([
  "text",
  "search",
  "email",
  "url",
  "tel",
  "number",
  "password"
]);

const EDITABLE_SELECTORS = [
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

let lastFocusedEditable = null;
let lastInteractedElement = null;
const DEBUG_PREFIX = "[SowisoHelper][content]";

function describeElement(element) {
  if (!element || !(element instanceof Element)) {
    return null;
  }

  return {
    tag: element.tagName,
    id: element.id || null,
    classes: element.className || null,
    role: element.getAttribute("role"),
    contenteditable: element.getAttribute("contenteditable"),
    type: element instanceof HTMLInputElement ? element.type : null
  };
}

function deepActiveElement(root = document) {
  let current = root.activeElement || null;

  while (current && current.shadowRoot && current.shadowRoot.activeElement) {
    current = current.shadowRoot.activeElement;
  }

  return current;
}

function normalizeEditable(target) {
  if (!target || !(target instanceof Element)) {
    return null;
  }

  if (target instanceof HTMLTextAreaElement) {
    return target;
  }

  if (target instanceof HTMLInputElement) {
    const inputType = (target.type || "text").toLowerCase();
    return ALLOWED_INPUT_TYPES.has(inputType) ? target : null;
  }

  if (target.isContentEditable) {
    return target;
  }

  const editableAncestor = target.closest('[contenteditable]:not([contenteditable="false"])');
  if (editableAncestor) {
    return editableAncestor;
  }

  const textRoleAncestor = target.closest("[role='textbox']");
  if (textRoleAncestor) {
    return textRoleAncestor;
  }

  return null;
}

function isConnected(element) {
  return Boolean(element && element.isConnected);
}

function isVisible(element) {
  if (!element || !(element instanceof HTMLElement)) {
    return false;
  }

  const rect = element.getBoundingClientRect();
  const style = window.getComputedStyle(element);
  return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
}

function isLikelySowisoMathTextarea(element) {
  if (!(element instanceof HTMLTextAreaElement)) {
    return false;
  }

  const className = (element.className || "").toLowerCase();
  return className.includes("math-editor") || className.includes("mathdoxformula");
}

function dispatchEditEvents(element, text = "") {
  try {
    element.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: text, inputType: "insertText" }));
  } catch (_error) {
    // Ignore environments where InputEvent construction is restricted.
  }

  element.dispatchEvent(new Event("input", { bubbles: true }));
  element.dispatchEvent(new Event("change", { bubbles: true }));
}

function insertIntoTextControl(element, text) {
  element.focus();
  const before = element.value;

  const start = element.selectionStart ?? element.value.length;
  const end = element.selectionEnd ?? start;

  if (typeof element.setRangeText === "function") {
    element.setRangeText(text, start, end, "end");
  } else {
    element.value = `${element.value.slice(0, start)}${text}${element.value.slice(end)}`;
  }

  dispatchEditEvents(element, text);
  return {
    changed: element.value !== before,
    beforeLength: before.length,
    afterLength: element.value.length
  };
}

function insertIntoContentEditable(element, text) {
  element.focus();

  let usedExecCommand = false;

  try {
    usedExecCommand = document.execCommand("insertText", false, text);
  } catch (_error) {
    usedExecCommand = false;
  }

  if (usedExecCommand) {
    dispatchEditEvents(element, text);
    return true;
  }

  const selection = window.getSelection();

  if (!selection || selection.rangeCount === 0) {
    element.append(document.createTextNode(text));
    dispatchEditEvents(element, text);
    return true;
  }

  const range = selection.getRangeAt(0);
  range.deleteContents();

  const node = document.createTextNode(text);
  range.insertNode(node);

  range.setStartAfter(node);
  range.setEndAfter(node);
  selection.removeAllRanges();
  selection.addRange(range);

  dispatchEditEvents(element, text);
  return true;
}

function tryDispatchPaste(target, text) {
  try {
    const data = new DataTransfer();
    data.setData("text/plain", text);

    const pasteEvent = new ClipboardEvent("paste", {
      clipboardData: data,
      bubbles: true,
      cancelable: true
    });

    const dispatched = target.dispatchEvent(pasteEvent);

    // If default was prevented, an editor likely consumed the paste event.
    if (!dispatched || pasteEvent.defaultPrevented) {
      return true;
    }
  } catch (_error) {
    // Ignore and continue with other strategies.
  }

  return false;
}

function queryEditableInside(element) {
  if (!element || !(element instanceof Element)) {
    return null;
  }

  const nested = element.querySelector(EDITABLE_SELECTORS.join(", "));
  return normalizeEditable(nested);
}

function resolveInsertionTarget() {
  if (isConnected(lastFocusedEditable)) {
    return normalizeEditable(lastFocusedEditable);
  }

  const active = normalizeEditable(deepActiveElement(document));
  if (active) {
    return active;
  }

  if (isConnected(lastInteractedElement)) {
    const direct = normalizeEditable(lastInteractedElement);
    if (direct) {
      return direct;
    }

    const nested = queryEditableInside(lastInteractedElement);
    if (nested) {
      return nested;
    }
  }

  const candidates = [...document.querySelectorAll(EDITABLE_SELECTORS.join(","))]
    .map((element) => normalizeEditable(element))
    .filter(Boolean);

  const visible = candidates.find((candidate) => isVisible(candidate));
  return visible || candidates[0] || null;
}

function focusAndResolveFromElement(element) {
  if (!element || !(element instanceof Element)) {
    return null;
  }

  try {
    element.focus();
    element.click();
  } catch (_error) {
    // Keep going.
  }

  const active = normalizeEditable(deepActiveElement(document));
  if (active) {
    return active;
  }

  return queryEditableInside(element);
}

function insertFormula(formula) {
  let target = resolveInsertionTarget();
  const debug = {
    activeElement: describeElement(document.activeElement),
    lastFocused: describeElement(lastFocusedEditable),
    lastInteracted: describeElement(lastInteractedElement)
  };

  if (!target && lastInteractedElement) {
    target = focusAndResolveFromElement(lastInteractedElement);
  }

  if (!target) {
    return { ok: false, error: "No editable input found. Click into an answer field first.", debug };
  }

  debug.target = describeElement(target);

  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
    if (isLikelySowisoMathTextarea(target)) {
      debug.method = "text-control";
      debug.mathTextareaHidden = !isVisible(target);
      return { ok: false, error: "Detected hidden MathDox textarea. Click the pink input slot and use keyboard insertion.", debug };
    }

    const textResult = insertIntoTextControl(target, formula);
    debug.method = "text-control";
    debug.textResult = textResult;

    if (!textResult.changed) {
      return { ok: false, error: "Text target did not change.", debug };
    }

    return { ok: true, debug };
  }

  if (target.isContentEditable) {
    insertIntoContentEditable(target, formula);
    debug.method = "contenteditable";
    return { ok: true, debug };
  }

  if (tryDispatchPaste(target, formula)) {
    debug.method = "paste-event";
    return { ok: true, debug };
  }

  const focused = focusAndResolveFromElement(target);
  debug.focusedAfterFallback = describeElement(focused);

  if (focused instanceof HTMLInputElement || focused instanceof HTMLTextAreaElement) {
    if (isLikelySowisoMathTextarea(focused)) {
      debug.method = "focused-text-control";
      debug.mathTextareaHidden = !isVisible(focused);
      return { ok: false, error: "Focused target is hidden MathDox textarea. Click the pink input slot and use keyboard insertion.", debug };
    }

    const textResult = insertIntoTextControl(focused, formula);
    debug.method = "focused-text-control";
    debug.textResult = textResult;

    if (!textResult.changed) {
      return { ok: false, error: "Focused text target did not change.", debug };
    }

    return { ok: true, debug };
  }

  if (focused && focused.isContentEditable) {
    insertIntoContentEditable(focused, formula);
    debug.method = "focused-contenteditable";
    return { ok: true, debug };
  }

  return { ok: false, error: "Found an editor container but could not insert text.", debug };
}

document.addEventListener(
  "focusin",
  (event) => {
    const editable = normalizeEditable(event.target);
    if (editable) {
      lastFocusedEditable = editable;
    }
  },
  true
);

document.addEventListener(
  "pointerdown",
  (event) => {
    if (event.target instanceof Element) {
      lastInteractedElement = event.target;
    }
  },
  true
);

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || message.type !== "INSERT_LATEX") {
    return;
  }

  const formula = typeof message.formula === "string" ? message.formula : "";
  console.log(DEBUG_PREFIX, "Insert message received", { formulaLength: formula.length });

  if (!formula) {
    const result = { ok: false, error: "No formula was provided." };
    console.log(DEBUG_PREFIX, "Insert result", result);
    sendResponse(result);
    return;
  }

  const result = insertFormula(formula);
  console.log(DEBUG_PREFIX, "Insert result", result);
  sendResponse(result);
});
})();
