"""
Clipboard Inspector — Dumps all available clipboard formats.
Usage:
    1. Copy a nested list FROM Slack (select items, Ctrl+C)
    2. Run: python clipboard_inspect.py
    3. Send me the output
"""

import ctypes
from ctypes import wintypes

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Set up proper 64-bit types
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
kernel32.GlobalSize.restype = ctypes.c_size_t
kernel32.GlobalSize.argtypes = [ctypes.c_void_p]
user32.GetClipboardData.restype = ctypes.c_void_p
user32.GetClipboardData.argtypes = [wintypes.UINT]

# Standard format names
STANDARD_FORMATS = {
    1: "CF_TEXT", 2: "CF_BITMAP", 3: "CF_METAFILEPICT",
    4: "CF_SYLK", 5: "CF_DIF", 6: "CF_TIFF",
    7: "CF_OEMTEXT", 8: "CF_DIB", 9: "CF_PALETTE",
    10: "CF_PENDATA", 11: "CF_RIFF", 12: "CF_WAVE",
    13: "CF_UNICODETEXT", 14: "CF_ENHMETAFILE",
    15: "CF_HDROP", 16: "CF_LOCALE", 17: "CF_DIBV5",
}


def get_format_name(fmt):
    if fmt in STANDARD_FORMATS:
        return STANDARD_FORMATS[fmt]
    buf = ctypes.create_unicode_buffer(256)
    if user32.GetClipboardFormatNameW(fmt, buf, 256):
        return buf.value
    return "Unknown({})".format(fmt)


def get_clipboard_data(fmt):
    hMem = user32.GetClipboardData(fmt)
    if not hMem:
        return None
    size = kernel32.GlobalSize(hMem)
    pMem = kernel32.GlobalLock(hMem)
    if not pMem:
        return None
    try:
        data = ctypes.string_at(pMem, size)
        return data
    finally:
        kernel32.GlobalUnlock(hMem)


def main():
    if not user32.OpenClipboard(0):
        print("ERROR: Could not open clipboard")
        return

    try:
        # Enumerate all formats
        fmt = 0
        formats = []
        while True:
            fmt = user32.EnumClipboardFormats(fmt)
            if fmt == 0:
                break
            formats.append(fmt)

        print("=" * 70)
        print("CLIPBOARD FORMATS ({} total)".format(len(formats)))
        print("=" * 70)

        for fmt in formats:
            name = get_format_name(fmt)
            print("\n--- {} (id={}) ---".format(name, fmt))

            data = get_clipboard_data(fmt)
            if data is None:
                print("  [could not read]")
                continue

            print("  Size: {} bytes".format(len(data)))

            # For text-like formats, show content
            if name in ("CF_TEXT", "CF_OEMTEXT"):
                try:
                    text = data.decode('ascii', errors='replace').rstrip('\x00')
                    print("  Content:\n{}".format(text))
                except:
                    print("  [decode error]")

            elif name == "CF_UNICODETEXT":
                try:
                    text = data.decode('utf-16-le', errors='replace').rstrip('\x00')
                    print("  Content:\n{}".format(text))
                except:
                    print("  [decode error]")

            elif name == "HTML Format":
                try:
                    text = data.decode('utf-8', errors='replace').rstrip('\x00')
                    print("  Content:\n{}".format(text))
                except:
                    print("  [decode error]")

            else:
                # Show first 200 bytes as hex + ascii
                preview = data[:200]
                # Try UTF-8 decode
                try:
                    text = preview.decode('utf-8', errors='replace')
                    if text.isprintable() or '\n' in text or '\r' in text:
                        print("  Preview (text):\n{}".format(text))
                    else:
                        print("  Preview (hex): {}".format(preview.hex()[:200]))
                except:
                    print("  Preview (hex): {}".format(preview.hex()[:200]))

    finally:
        user32.CloseClipboard()


if __name__ == '__main__':
    main()
