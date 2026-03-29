# Sowieso LaTeX Helper (Chrome Extension)

> This project is still a WIP – More updates coming soon!

Small Chrome extension for `cloud.sowiso.nl` that lets you:

- Paste a LaTeX formula
- Preview it live
- Insert it into the active exercise input with one click

## Current MVP behavior

- The extension opens as a Chrome side panel (right sidebar).
- The panel has a formula textbox and render preview.
- You can choose insert mode:
  - `Raw LaTeX`
  - `Inline ($...$)`
  - `Block ($$...$$)`
- Click **Insert into page** or click the preview itself to insert.
- Theme switcher included: `Auto`, `Light`, `Dark` (saved locally).
- It inserts into the last focused input/textarea/contenteditable field on `cloud.sowiso.nl`.

## Install (unpacked)

1. Open Chrome and go to `chrome://extensions`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select this folder: `chrome-extension`.
5. Pin the extension to your toolbar.

## Usage

1. Open an exercise page on `https://cloud.sowiso.nl`.
2. Click inside the answer field once.
3. Click the extension icon to open the side panel.
4. Paste your LaTeX formula.
5. Click **Insert into page**.
6. If it does not insert, click the answer field once again and retry (the extension now retries across all iframes).

## Notes

- Preview currently uses the CodeCogs render endpoint over the network.
- If Sowiso expects a special formula format instead of raw LaTeX, we can add a conversion layer next.
