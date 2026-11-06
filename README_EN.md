# Outlook Image Helper & OCR

An NVDA screen reader add-on providing full accessibility, local optical character recognition (OCR), interactive multi-turn AI chat and descriptions (**Google Gemini**), prompt presets library, hyperlink activation, and clean batch saving for images and screenshots embedded in **Microsoft Outlook Classic** messages.

---

## ✨ Key Features

1. **Context Action Menu on Image List:**
   - Press `Applications`, `Shift+F10`, or `Alt+C` on any image in the list to open the context action menu without tabbing.
   - When multiple images are selected with `Space`, the menu automatically adapts to batch actions (batch OCR, batch AI, batch save).

2. **Prompt Presets Library & Custom Prompts:**
   - Choose from pre-configured AI instruction templates (*HR / Leave balances*, *App window / errors*, *Charts*, *Verbatim transcription*, etc.).
   - Create, edit, and delete custom prompt templates.

3. **Interactive Multi-turn AI Chat (Q&A):**
   - Follow-up questions about the image with full conversation and image context preserved (`Ctrl+Enter` sends question).

4. **Universal Rich Copy & Export:**
   - Copy individual answers or entire transcripts.
   - Option to include original image in Rich HTML (Base64) format. Pasting into Outlook or Word embeds the image alongside styled text.
   - Standalone export to `.html` or `.md` (Markdown).

5. **Hyperlink Detection & Activation:**
   - Detects embedded links (YouTube, LinkedIn) and enables opening them directly with `Ctrl+Enter`.

6. **Preceding Context & Clean Filenames (GDPR):**
   - Automatically extracts employee names or headers written directly above screenshots.
   - Clean local batch saving (e.g. `01_John_Doe.png`) runs 100% locally to protect privacy and comply with GDPR.

---

## ⌨️ Shortcuts

- `Ctrl+NVDA+A` (or `NVDA+Shift+G`, `Alt+Shift+G`) – Open image list in message.
- `NVDA+Windows+G` or `Alt+Shift+O` – Quick OCR on image.
- `NVDA+Windows+Shift+G` or `Alt+Shift+A` – Quick AI description.
- Inside the image list:
  - `Applications` / `Shift+F10` / `Alt+C` – Context action menu.
  - `Ctrl+Enter` – Open associated hyperlink in browser.
  - `Space` / `Shift+Arrows` – Multi-selection.

---

## 🛠️ Building the Add-on

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
```
