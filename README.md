# Sublime: Tab Note Converter — Sublime Text Plugin

Convert tab-indented notes into **Markdown**, **Rich Text**, or **Slack Message** format with a single right-click.

The Slack export produces properly nested lists - something widely considered impossible when pasting from external tools. [Here's how we figured it out.](#the-slack-rabbit-hole)

![Sublime Text 4](https://img.shields.io/badge/Sublime_Text-4-orange)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![No Dependencies](https://img.shields.io/badge/dependencies-none-green)

## AI use disclaimer

This tool was developed with heavy usage of Claude AI's Opus 4.6.

## What it does

You write quick notes in Sublime using tab indentation:

```
Legendary Weapons UI Review
	Creation UI
		move "forge" as 4th Trait
			Jon: Players should be able to see the create UI
			4th trait is more powerful, but still a trait
				not an E-spell
				probably also not the W-spell
		show requirements in Details UI
			Legendary Rating of X
	Encyclopedia
		Not so much a ranking, but a story-telling device
```

Then right-click → **Convert tab note to** → pick your target:

| Format            | What you get | Best for |
|-------------------|-------------|----------|
| **Markdown**      | Standard markdown with `#`, `**bold**`, `* bullets` | Confluence Pages, GitHub, anywhere that renders MD |
| **Rich Text**     | HTML on clipboard via `CF_HTML` | Confluence Live Docs, Word, Google Docs |
| **Slack Message** | Slack's proprietary `data-stringify` HTML via Chromium MIME | Slack - with fully nested lists |


## Input format

The tab depth determines what each line becomes:

| Tabs | Role | Output                   |
|------|------|--------------------------|
| 0 | Title / Heading | `# Heading`              |
| 1 | Section header | `**Bold header**`        |
| 2 | Top-level bullet | `* Bullet`               |
| 3 | Nested bullet | `  * Sub-bullet`         |
| 4+ | Deeper nesting | `    * Sub-sub-bullet` … |

Blank lines between sections are preserved. If you select text, only the selection is converted. Otherwise the entire buffer is used.


## Installation

Drop the `Tab Note Converter` folder into your Sublime Text Packages directory:

**Windows:** `%APPDATA%\Sublime Text[ 3]\Packages\Tab Note Converter\`

That's it. No dependencies, no external tools, no build step.


## Usage

**Context menu:** Right-click → Convert tab note to → Markdown / Rich Text / Slack Message

**Command Palette:** `Ctrl+Shift+P` → `Tab Note Converter: Markdown` / `Tab Note Converter: Rich Text` / `Tab Note Converter: Slack Message`


## Platform support

Currently **Windows only**. The Rich Text and Slack exports use the Windows clipboard API (`ctypes` → `user32.dll` / `kernel32.dll`) to set multiple clipboard formats simultaneously. A PowerShell fallback exists for Rich Text if ctypes misbehaves.

macOS and Linux contributions are welcome - the conversion logic is platform-independent, only the clipboard writing needs platform-specific code.


---

## The Slack rabbit hole

This section is for anyone who's ever tried to paste nested lists into Slack and hit a wall. It's also the story of how this plugin went from "quick weekend project" to "reverse-engineering proprietary clipboard formats."

### The problem

Slack flattens nested lists on paste. Always. Every rich text editor, every HTML clipboard format, every markdown trick - Slack turns your carefully indented outline into a flat list of bullets. The internet consensus: *it's not possible.*

But there's one exception: **copying nested lists from within Slack itself preserves nesting.** If Slack can do it to itself, we should be able to figure out what it's doing.

### The investigation

Step 1 was writing a clipboard inspector script (included in this repo as `tools/clipboard_inspect.py`) that dumps every format on the Windows clipboard. We used it to compare what Slack puts on the clipboard versus what we were generating.

The obvious suspect was `CF_HTML` - the standard Windows clipboard format for rich text. Our HTML was structurally correct (proper nested `<ul><li>` elements) and even used Slack's own inline CSS styles. But Slack ignored it completely.

Then we noticed something in the clipboard dump that most people overlook:

```
--- Chromium Web Custom MIME Data Format (id=49834) ---
  Size: 11224 bytes
  Preview (hex): d42b0000010000000a000000730006c00...
```

### Chromium Web Custom MIME Data Format

Slack is an Electron app, which means it runs on Chromium. Chromium has a mechanism for passing custom MIME types through the OS clipboard: the **Chromium Web Custom MIME Data Format**. It uses Chromium's internal `Pickle` binary serializer to store key-value pairs of `(mime_type, content)`.

We wrote a decoder and found that Slack registers a custom MIME type: `slack/html`. When pasting, Slack checks for this MIME type *first*, before falling back to `CF_HTML`.

### The format

The `slack/html` content is completely different from the `CF_HTML` content. Where `CF_HTML` uses inline CSS styles, `slack/html` uses Slack's internal CSS classes and `data-stringify-*` attributes:

```html
<!-- CF_HTML (what we tried first - Slack ignores this for nesting) -->
<ul style="margin: 4px 0 4px 24px; padding: 0; list-style-position: outside">
  <li style="margin: 2px 0; padding: 0; color: #1d1c1d">item</li>
</ul>

<!-- slack/html (what actually works) -->
<ul data-stringify-type="unordered-list" data-list-tree="true"
    class="p-rich_text_list p-rich_text_list__bullet p-rich_text_list--nested"
    data-indent="0" data-border="0">
  <li data-stringify-indent="0" data-stringify-border="0">item</li>
</ul>
```

Key differences:
- `data-indent` on `<ul>` and `data-stringify-indent` on `<li>` control nesting depth
- `data-stringify-type="unordered-list"` identifies the list type
- `data-list-tree="true"` enables the tree structure
- `<b data-stringify-type="bold">` instead of `<b style="font-weight: 700">`
- `<meta charset="utf-8">` prefix required
- Everything minified - no whitespace between tags
- Double quotes in content are NOT HTML-escaped

### The Pickle format

The Chromium Pickle serialization is straightforward (little-endian):

```
uint32  payload_size
uint32  num_entries
For each entry:
    uint32  mime_type_char_count
    char16  mime_type[]          (UTF-16LE)
    [padding to 4-byte alignment]
    uint32  content_char_count
    char16  content[]            (UTF-16LE)
    [padding to 4-byte alignment]
```

### The clipboard sandwich

The final "Convert for Slack" command sets **three** formats in a single clipboard session:

1. **Chromium Web Custom MIME Data Format** - contains `slack/html` with the proprietary markup. Slack reads this first.
2. **CF_HTML** - same content with inline styles, for apps that read standard HTML (Confluence, Word, etc.)
3. **CF_UNICODETEXT** - Markdown plain text as a final fallback.

This way, the same clipboard paste works in Slack (nested lists!), Confluence (rich text), and any plain text field (readable Markdown).

### Tools

The `tools/` directory contains the clipboard inspection scripts used during development:

- `clipboard_inspect.py` - dumps all clipboard formats
- `clipboard_inspect_v2.py` - fully decodes Chromium Web Custom MIME data

These are useful for debugging clipboard issues or reverse-engineering other Electron apps' clipboard formats.


---

## License

MIT


## Credits

Built in a wild late-night pair-programming session between a human and an AI, fueled by Big Game Dev Energy and the refusal to accept "not possible" as an answer.
