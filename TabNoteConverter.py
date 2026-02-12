"""
Tab Note Converter - Sublime Text Plugin
Converts raw tab-indented meeting notes to Markdown or Rich Text
and copies to clipboard.

Raw format (tab-indented):
    Title Line              (0 tabs → # Heading)
        Section Header      (1 tab  → **Bold section**)
            Bullet point    (2 tabs → * Bullet)
                Nested      (3 tabs →   * Nested bullet)

Install: Drop this folder into your Packages directory.
Usage:   Right-click → Convert tab note to → Markdown / Rich Text
         Or via Command Palette: Tab Note Converter: Markdown / Rich Text
"""

import sublime
import sublime_plugin
import re
import sys


# ---------------------------------------------------------------------------
# Raw → Markdown conversion
# ---------------------------------------------------------------------------

def raw_to_markdown(text):
    """Convert raw tab-indented notes into Markdown.

    Indentation mapping:
        0 tabs → # Heading
        1 tab  → **Bold section header**
        2+ tabs → Bullet points, nesting depth = tabs - 2
    """
    lines = text.split('\n')
    result = []

    for line in lines:
        stripped = line.rstrip()

        # Blank line → preserve as separator
        if not stripped:
            result.append('')
            continue

        # Count leading tabs
        tab_count = len(line) - len(line.lstrip('\t'))
        content = stripped.lstrip('\t')

        if tab_count == 0:
            # Top-level heading
            result.append('# {}'.format(content))
        elif tab_count == 1:
            # Section header (bold)
            result.append('')
            result.append('**{}**'.format(content))
        else:
            # Bullet points: 2 tabs = *, 3 tabs = "  *", etc.
            indent = '  ' * (tab_count - 2)
            result.append('{}* {}'.format(indent, content))

    # Collapse multiple consecutive blank lines into one
    output = '\n'.join(result)
    output = re.sub(r'\n{3,}', '\n\n', output)
    return output


# ---------------------------------------------------------------------------
# Raw → Slack-flavored HTML conversion
# ---------------------------------------------------------------------------

# Slack's internal HTML templates (reverse-engineered from clipboard output).
# Slack uses data-stringify-* attributes and specific CSS classes for its
# rich text editor to recognize pasted content as structured data.

_SLACK_META = '<meta charset="utf-8">'
_SLACK_DIV_OPEN = '<div class="p-rich_text_section">'
_SLACK_DIV_CLOSE = '</div>'
_SLACK_BOLD = '<b data-stringify-type="bold">{}</b>'
_SLACK_BR = '<br aria-hidden="true">'
_SLACK_PARA_BREAK = '<span aria-label="&nbsp;" class="c-mrkdwn__br" data-stringify-type="paragraph-break"></span>'

def _slack_ul_open(indent):
    return (
        '<ul data-stringify-type="unordered-list" data-list-tree="true" '
        'class="p-rich_text_list p-rich_text_list__bullet '
        'p-rich_text_list--nested" data-indent="{}" '
        'data-border="0">'.format(indent)
    )

def _slack_li_open(indent):
    return '<li data-stringify-indent="{}" data-stringify-border="0">'.format(indent)


def _slack_escape(text):
    """Escape HTML but NOT quotes — Slack doesn't escape them."""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;'))


def raw_to_slack_html(text):
    """Convert raw tab-indented notes into Slack's internal HTML format.

    Slack uses a proprietary HTML format with data-stringify-* attributes
    and specific CSS classes. This is stored in the 'slack/html' MIME type
    inside Chromium's Web Custom MIME Data Format clipboard entry.

    Indentation mapping:
        0 tabs → Bold heading in a section div
        1 tab  → Bold section header in a section div
        2+ tabs → Nested <ul><li> with data-indent attributes
    """
    lines = text.split('\n')
    parts = [_SLACK_META]
    list_stack = []  # tracks indent levels of open lists
    div_open = False

    def _close_all_lists():
        nonlocal list_stack
        while list_stack:
            parts.append('</li></ul>')
            list_stack.pop()

    def _ensure_div():
        nonlocal div_open
        if not div_open:
            parts.append(_SLACK_DIV_OPEN)
            div_open = True

    def _close_div():
        nonlocal div_open
        if div_open:
            parts.append(_SLACK_DIV_CLOSE)
            div_open = False

    first_heading = True

    for line in lines:
        stripped = line.rstrip()

        if not stripped:
            _close_all_lists()
            continue

        tab_count = len(line) - len(line.lstrip('\t'))
        content = _slack_escape(stripped.lstrip('\t'))

        if tab_count <= 1:
            # Header line — close any open lists first
            _close_all_lists()

            if tab_count == 0:
                # Top-level heading → UPPERCASE bold
                if not div_open and not first_heading:
                    # Add paragraph break between sections
                    _ensure_div()
                    parts.append(_SLACK_PARA_BREAK)
                else:
                    _ensure_div()
                first_heading = False
                parts.append(_SLACK_BOLD.format(content.upper()))
            else:
                # Section header → bold
                _ensure_div()
                parts.append(_SLACK_PARA_BREAK)
                parts.append(_SLACK_BOLD.format(content))
                parts.append(_SLACK_BR)
                _close_div()
        else:
            # Bullet point
            if div_open:
                _close_div()

            depth = tab_count - 2  # 0-indexed indent level

            if not list_stack:
                # Start first list
                parts.append(_slack_ul_open(depth))
                parts.append(_slack_li_open(depth) + content)
                list_stack.append(depth)
            elif depth > list_stack[-1]:
                # Going deeper — open nested list(s)
                while depth > list_stack[-1]:
                    next_depth = list_stack[-1] + 1
                    parts.append(_slack_ul_open(next_depth))
                    if next_depth == depth:
                        parts.append(_slack_li_open(depth) + content)
                    else:
                        parts.append(_slack_li_open(next_depth))
                    list_stack.append(next_depth)
            elif depth == list_stack[-1]:
                # Same level
                parts.append('</li>')
                parts.append(_slack_li_open(depth) + content)
            else:
                # Going shallower
                while list_stack and list_stack[-1] > depth:
                    parts.append('</li></ul>')
                    list_stack.pop()
                parts.append('</li>')
                parts.append(_slack_li_open(depth) + content)

    # Close remaining open elements
    _close_all_lists()
    _close_div()

    return ''.join(parts)


#
# NOTE: This intentionally does NOT handle arbitrary markdown. It only
# converts the limited subset produced by raw_to_markdown(). If this plugin
# ever needs to handle arbitrary markdown, swap this out for `mistune`.
# ---------------------------------------------------------------------------

def _escape_html(text):
    """Escape HTML special characters."""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def _inline_format(text):
    """Handle inline markdown formatting (bold, italic, inline code)."""
    escaped = _escape_html(text)
    # Bold: **text**
    escaped = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', escaped)
    # Italic: *text* (but not inside bold)
    escaped = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'<em>\1</em>', escaped)
    # Inline code: `text`
    escaped = re.sub(r'`(.+?)`', r'<code>\1</code>', escaped)
    return escaped


def markdown_to_html(md):
    """Convert our generated Markdown subset to HTML."""
    lines = md.split('\n')
    html_parts = []
    current_depth = 0  # 0 = not in a list

    for line in lines:
        stripped = line.strip()

        # Blank line → close any open lists
        if not stripped:
            for _ in range(current_depth):
                html_parts.append('</li></ul>')
            current_depth = 0
            continue

        # Heading: # text
        heading_match = re.match(r'^(#{1,6})\s+(.+)$', stripped)
        if heading_match:
            for _ in range(current_depth):
                html_parts.append('</li></ul>')
            current_depth = 0
            level = len(heading_match.group(1))
            content = _inline_format(heading_match.group(2))
            html_parts.append('<h{0}>{1}</h{0}>'.format(level, content))
            continue

        # Standalone bold line (section header): **text**
        bold_match = re.match(r'^\*\*(.+)\*\*$', stripped)
        if bold_match:
            for _ in range(current_depth):
                html_parts.append('</li></ul>')
            current_depth = 0
            content = _escape_html(bold_match.group(1))
            html_parts.append('<p><strong>{}</strong></p>'.format(content))
            continue

        # Bullet point: (spaces)* text
        bullet_match = re.match(r'^(\s*)\*\s+(.+)$', line)
        if bullet_match:
            indent_spaces = len(bullet_match.group(1))
            depth = indent_spaces // 2 + 1
            content = _inline_format(bullet_match.group(2))

            if depth > current_depth:
                # Going deeper — open new list levels
                for _ in range(depth - current_depth):
                    html_parts.append('<ul><li>')
                # Attach content to the last opened <li>
                html_parts[-1] = '<ul><li>{}'.format(content)
            elif depth == current_depth:
                # Same level
                html_parts.append('</li><li>{}'.format(content))
            else:
                # Going shallower — close deeper levels, then new item
                for _ in range(current_depth - depth):
                    html_parts.append('</li></ul>')
                html_parts.append('</li><li>{}'.format(content))

            current_depth = depth
            continue

        # Fallback: plain paragraph
        for _ in range(current_depth):
            html_parts.append('</li></ul>')
        current_depth = 0
        html_parts.append('<p>{}</p>'.format(_inline_format(stripped)))

    # Close any remaining open lists
    for _ in range(current_depth):
        html_parts.append('</li></ul>')

    return '\n'.join(html_parts)


# ---------------------------------------------------------------------------
# CF_HTML payload builder
# ---------------------------------------------------------------------------

def _build_cf_html(html_fragment):
    """Build a CF_HTML formatted payload for the Windows clipboard.

    CF_HTML requires a UTF-8 encoded header with byte offsets pointing to
    the HTML and fragment boundaries. See:
    https://docs.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
    """
    header_template = (
        "Version:0.9\r\n"
        "StartHTML:{:010d}\r\n"
        "EndHTML:{:010d}\r\n"
        "StartFragment:{:010d}\r\n"
        "EndFragment:{:010d}\r\n"
    )

    prefix = "<html><body>\r\n<!--StartFragment-->"
    suffix = "<!--EndFragment-->\r\n</body></html>"

    # Calculate byte offsets (must be byte-level, not char-level)
    header_len = len(header_template.format(0, 0, 0, 0).encode('utf-8'))
    start_html = header_len
    start_fragment = start_html + len(prefix.encode('utf-8'))
    end_fragment = start_fragment + len(html_fragment.encode('utf-8'))
    end_html = end_fragment + len(suffix.encode('utf-8'))

    header = header_template.format(start_html, end_html, start_fragment, end_fragment)
    return header + prefix + html_fragment + suffix


# ---------------------------------------------------------------------------
# Chromium Web Custom MIME Data Format builder
# ---------------------------------------------------------------------------

def _build_chromium_custom_mime(entries):
    """Build a Chromium Web Custom MIME Data Format binary payload.

    Chromium uses its Pickle serializer to store custom MIME types on the
    clipboard. Electron apps (like Slack) read this to get their own
    internal formats (e.g. 'slack/html').

    Format (little-endian):
        uint32: payload_size (size of everything after this field)
        uint32: num_entries
        For each entry (MIME type, content):
            uint32: string_char_count (UTF-16 chars)
            char16[]: string data (UTF-16LE)
            padding to 4-byte alignment

    Args:
        entries: list of (mime_type: str, content: str) tuples
    Returns:
        bytes: the complete binary payload
    """
    import struct

    def _pickle_write_string16(parts, text):
        """Append a Pickle-serialized UTF-16 string to parts list."""
        encoded = text.encode('utf-16-le')
        char_count = len(text)
        parts.append(struct.pack('<I', char_count))
        parts.append(encoded)
        # Pad to 4-byte alignment
        remainder = len(encoded) % 4
        if remainder:
            parts.append(b'\x00' * (4 - remainder))

    # Build the payload (everything after the header size field)
    payload_parts = []
    payload_parts.append(struct.pack('<I', len(entries)))

    for mime_type, content in entries:
        _pickle_write_string16(payload_parts, mime_type)
        _pickle_write_string16(payload_parts, content)

    payload = b''.join(payload_parts)

    # Prepend the payload size header
    return struct.pack('<I', len(payload)) + payload


# ---------------------------------------------------------------------------
# Clipboard: Rich Text (HTML) — Windows via ctypes, with PowerShell fallback
# ---------------------------------------------------------------------------

def _alloc_clipboard_data(kernel32, data_bytes):
    """Allocate a moveable global memory block and copy data into it."""
    import ctypes

    GMEM_MOVEABLE = 0x0002
    hMem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data_bytes) + 2)
    if not hMem:
        raise RuntimeError("GlobalAlloc failed")

    pMem = kernel32.GlobalLock(hMem)
    if not pMem:
        raise RuntimeError("GlobalLock failed")

    ctypes.memmove(pMem, data_bytes, len(data_bytes))
    # Null-terminate (2 bytes for UTF-16 safety)
    ctypes.memset(pMem + len(data_bytes), 0, 2)
    kernel32.GlobalUnlock(hMem)
    return hMem


def _set_clipboard_html_ctypes(html, plain_text, chromium_custom_data=None):
    """Put CF_HTML, CF_UNICODETEXT, and optionally Chromium custom MIME data
    onto the Windows clipboard in a single session.

    Both formats are set in a single clipboard session so that:
    - Apps reading CF_HTML (Confluence) get rich text
    - Apps reading CF_UNICODETEXT (Slack) get markdown plain text
    - Apps reading Chromium custom MIME (Slack) get slack/html with nesting
    """
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    # Critical: set correct return/arg types for 64-bit pointer safety.
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    user32.SetClipboardData.restype = ctypes.c_void_p
    user32.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]

    CF_UNICODETEXT = 13
    CF_HTML = user32.RegisterClipboardFormatW("HTML Format")
    if not CF_HTML:
        raise RuntimeError("Failed to register CF_HTML clipboard format")

    # Prepare payloads
    html_payload = _build_cf_html(html).encode('utf-8')
    text_payload = plain_text.encode('utf-16-le')

    if not user32.OpenClipboard(0):
        raise RuntimeError("Could not open clipboard")

    try:
        user32.EmptyClipboard()

        # Set CF_HTML (for Confluence and other HTML-aware apps)
        hHtml = _alloc_clipboard_data(kernel32, html_payload)
        if not user32.SetClipboardData(CF_HTML, hHtml):
            raise RuntimeError("SetClipboardData CF_HTML failed")

        # Set CF_UNICODETEXT (for plain-text-only apps)
        hText = _alloc_clipboard_data(kernel32, text_payload)
        if not user32.SetClipboardData(CF_UNICODETEXT, hText):
            raise RuntimeError("SetClipboardData CF_UNICODETEXT failed")

        # Optionally set Chromium Web Custom MIME Data Format
        if chromium_custom_data is not None:
            CF_CHROMIUM = user32.RegisterClipboardFormatW(
                "Chromium Web Custom MIME Data Format"
            )
            if CF_CHROMIUM:
                hChromium = _alloc_clipboard_data(kernel32, chromium_custom_data)
                if not user32.SetClipboardData(CF_CHROMIUM, hChromium):
                    print("[Tab Note Converter] Warning: SetClipboardData Chromium custom MIME failed")
    finally:
        user32.CloseClipboard()


def _set_clipboard_html_powershell(html, plain_text, chromium_custom_data=None):
    """Fallback: use PowerShell + .NET to set CF_HTML and plain text.

    NOTE: Chromium custom MIME data is not supported via PowerShell fallback.
    This means Slack nested lists won't work if ctypes fails.
    """
    import subprocess
    import tempfile
    import os

    payload = _build_cf_html(html)

    # Write CF_HTML payload as raw bytes to a temp file
    tmp_path = os.path.join(tempfile.gettempdir(), '_copyas_clipboard.bin')
    with open(tmp_path, 'wb') as f:
        f.write(payload.encode('utf-8'))

    # Escape backslashes for PowerShell string interpolation
    ps_path = tmp_path.replace('\\', '\\\\')
    # Escape single quotes in plain text for PowerShell
    ps_text = plain_text.replace("'", "''")

    ps_script = (
        'Add-Type -AssemblyName System.Windows.Forms; '
        '$bytes = [System.IO.File]::ReadAllBytes("{}"); '
        '$stream = New-Object System.IO.MemoryStream(,$bytes); '
        '$dataObj = New-Object System.Windows.Forms.DataObject; '
        '$dataObj.SetData("HTML Format", $stream); '
        "$dataObj.SetData([System.Windows.Forms.DataFormats]::UnicodeText, '{}'); "
        '[System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)'
    ).format(ps_path, ps_text)

    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
            capture_output=True,
            timeout=5,
            # CREATE_NO_WINDOW = 0x08000000 — prevents a console flash
            creationflags=0x08000000
        )
        if result.returncode != 0:
            stderr = result.stderr.decode('utf-8', errors='replace')
            raise RuntimeError("PowerShell clipboard failed: {}".format(stderr))
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def set_clipboard_html(html, plain_text, chromium_custom_data=None):
    """Copy HTML + plain text to clipboard as rich text.

    Strategy: ctypes (fast) → PowerShell fallback → raw HTML as plain text.
    """
    if sys.platform != 'win32':
        sublime.set_clipboard(html)
        sublime.status_message("Tab Note Converter: Rich text not supported on this OS — copied raw HTML")
        return False

    # Try ctypes first (fast, no subprocess overhead)
    try:
        _set_clipboard_html_ctypes(html, plain_text, chromium_custom_data)
        return True
    except Exception as e:
        print("[Tab Note Converter] ctypes clipboard failed ({}), trying PowerShell...".format(e))

    # Fallback: PowerShell + .NET (no Chromium custom MIME support)
    try:
        _set_clipboard_html_powershell(html, plain_text, chromium_custom_data)
        return True
    except Exception as e:
        print("[Tab Note Converter] PowerShell clipboard also failed: {}".format(e))

    # Last resort: plain HTML as text
    sublime.set_clipboard(html)
    sublime.status_message("Tab Note Converter: Could not set rich text — copied raw HTML (check console for errors)")
    return False


# ---------------------------------------------------------------------------
# Helper: get content to convert
# ---------------------------------------------------------------------------

def _get_content(view):
    """Return the selected text, or the entire buffer if nothing is selected."""
    sel = view.sel()
    if sel and len(sel) == 1 and not sel[0].empty():
        return view.substr(sel[0])
    return view.substr(sublime.Region(0, view.size()))


def _normalize_output(text):
    """Strip leading whitespace, ensure exactly one trailing newline."""
    return text.strip() + '\n'


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

class ConvertTabNoteToMarkdownCommand(sublime_plugin.TextCommand):
    """Convert raw notes to Markdown and copy to clipboard."""

    def run(self, edit):
        content = _get_content(self.view)
        markdown = _normalize_output(raw_to_markdown(content))
        sublime.set_clipboard(markdown)
        sublime.status_message("Converted to Markdown")

    def is_enabled(self):
        return self.view.size() > 0


class ConvertTabNoteToSlackMessageCommand(sublime_plugin.TextCommand):
    """Convert raw notes to Slack-compatible HTML and copy to clipboard.

    Sets three clipboard formats:
    - Chromium custom MIME with 'slack/html' (Slack reads this first → nested lists work)
    - CF_HTML (fallback for other apps)
    - CF_UNICODETEXT (plain text fallback)
    """

    def run(self, edit):
        content = _get_content(self.view)
        slack_html = raw_to_slack_html(content)
        markdown = _normalize_output(raw_to_markdown(content))

        # Build Chromium custom MIME payload with slack/html
        chromium_data = _build_chromium_custom_mime([
            ('slack/html', slack_html),
        ])

        success = set_clipboard_html(slack_html, markdown, chromium_data)
        if success:
            sublime.status_message("Converted for Slack")

    def is_enabled(self):
        return self.view.size() > 0


class ConvertTabNoteToRichtextCommand(sublime_plugin.TextCommand):
    """Convert raw notes to Rich Text (HTML) and copy to clipboard."""

    def run(self, edit):
        content = _get_content(self.view)
        markdown = raw_to_markdown(content)
        html = markdown_to_html(markdown)
        plain_text = _normalize_output(markdown)
        success = set_clipboard_html(html, plain_text)
        if success:
            sublime.status_message("Converted to Rich Text")

    def is_enabled(self):
        return self.view.size() > 0
