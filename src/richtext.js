// Minimal contenteditable rich-text support: bold, bullet list, numbered
// list. Uses the classic execCommand API — deprecated but still universally
// supported for exactly this kind of basic toggle, and avoids pulling in a
// full editor library for three buttons.

function isWholeLineBold(node) {
  return (
    node.children.length === 1 &&
    ["B", "STRONG"].includes(node.children[0].tagName) &&
    node.children[0].textContent === node.textContent
  );
}

function blockFromListItem(li, type) {
  const text = li.textContent.trim();
  return text ? { type, bold: isWholeLineBold(li), text } : null;
}

function blockFromContainer(node) {
  const text = node.textContent.trim();
  return text ? { type: "paragraph", bold: isWholeLineBold(node), text } : null;
}

// Parses a contenteditable's content into an ordered list of
// { type: "paragraph" | "bullet" | "number", bold, text } blocks.
//
// execCommand's DOM output isn't reliably flat — a list can end up nested
// inside a wrapping <div> or <b> rather than sitting as a direct child of
// the editor — so this walks recursively instead of assuming one level.
export function parseRichText(root) {
  const blocks = [];
  const state = { buffer: "" };

  const flush = () => {
    const text = state.buffer.trim();
    if (text) blocks.push({ type: "paragraph", bold: false, text });
    state.buffer = "";
  };

  const walk = (container) => {
    for (const node of container.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        state.buffer += node.textContent;
      } else if (node.nodeName === "BR") {
        flush();
      } else if (node.nodeName === "UL" || node.nodeName === "OL") {
        flush();
        const type = node.nodeName === "UL" ? "bullet" : "number";
        for (const li of node.querySelectorAll(":scope > li")) {
          const block = blockFromListItem(li, type);
          if (block) blocks.push(block);
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.querySelector("ul, ol")) {
          flush();
          walk(node);
        } else {
          flush();
          const block = blockFromContainer(node);
          if (block) blocks.push(block);
        }
      }
    }
  };

  walk(root);
  flush();
  return blocks;
}

const COMMANDS = {
  bold: "bold",
  bullet: "insertUnorderedList",
  number: "insertOrderedList",
};

export function initRichTextEditors(root = document) {
  try {
    document.execCommand("defaultParagraphSeparator", false, "div");
  } catch (err) {
    // ignore — Enter will just fall back to whatever the browser does
  }

  root.querySelectorAll(".richtext-toolbar").forEach((toolbar) => {
    const targetId = toolbar.dataset.target;
    const editor = document.getElementById(targetId);
    if (!editor) return;

    toolbar.querySelectorAll("button[data-cmd]").forEach((btn) => {
      btn.addEventListener("click", () => {
        editor.focus();
        document.execCommand(COMMANDS[btn.dataset.cmd], false, null);
        updateToolbarState(toolbar, editor);
      });
    });

    const sync = () => updateToolbarState(toolbar, editor);
    editor.addEventListener("keyup", sync);
    editor.addEventListener("mouseup", sync);
    editor.addEventListener("focus", sync);
  });
}

function updateToolbarState(toolbar, editor) {
  toolbar.querySelectorAll("button[data-cmd]").forEach((btn) => {
    const active = document.queryCommandState(COMMANDS[btn.dataset.cmd]);
    btn.classList.toggle("active", active);
  });
}
