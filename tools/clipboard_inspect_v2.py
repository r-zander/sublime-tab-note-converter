"""
Clipboard Inspector v2 — Fully decodes Chromium Web Custom MIME Data Format.
Usage:
    1. Copy a nested list FROM Slack (select items, Ctrl+C)
    2. Run: python clipboard_inspect_v2.py
    3. Send me the output (or pipe to file: python clipboard_inspect_v2.py > dump.txt)
"""

import ctypes
from ctypes import wintypes
import struct

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
kernel32.GlobalSize.restype = ctypes.c_size_t
kernel32.GlobalSize.argtypes = [ctypes.c_void_p]
user32.GetClipboardData.restype = ctypes.c_void_p
user32.GetClipboardData.argtypes = [wintypes.UINT]


def get_clipboard_data(fmt):
    hMem = user32.GetClipboardData(fmt)
    if not hMem:
        return None
    size = kernel32.GlobalSize(hMem)
    pMem = kernel32.GlobalLock(hMem)
    if not pMem:
        return None
    try:
        return ctypes.string_at(pMem, size)
    finally:
        kernel32.GlobalUnlock(hMem)


def decode_chromium_pickle(data):
    """Decode Chromium's Pickle-serialized custom MIME data."""
    offset = 0

    # Header: uint32 payload_size
    payload_size = struct.unpack_from('<I', data, offset)[0]
    offset += 4
    print("  Payload size: {} bytes".format(payload_size))

    # uint32 num_entries
    num_entries = struct.unpack_from('<I', data, offset)[0]
    offset += 4
    print("  Number of entries: {}".format(num_entries))

    for i in range(num_entries):
        # Read MIME type string (UTF-16LE)
        char_count = struct.unpack_from('<I', data, offset)[0]
        offset += 4
        mime_bytes = data[offset:offset + char_count * 2]
        mime = mime_bytes.decode('utf-16-le')
        offset += char_count * 2
        # Alignment padding
        if (char_count * 2) % 4:
            offset += 4 - (char_count * 2) % 4

        # Read content string (UTF-16LE)
        char_count2 = struct.unpack_from('<I', data, offset)[0]
        offset += 4
        content_bytes = data[offset:offset + char_count2 * 2]
        content = content_bytes.decode('utf-16-le')
        offset += char_count2 * 2
        if (char_count2 * 2) % 4:
            offset += 4 - (char_count2 * 2) % 4

        print("\n  --- Entry {} ---".format(i))
        print("  MIME type: '{}'".format(mime))
        print("  Content length: {} chars".format(len(content)))
        print("  Content:")
        print(content)

    return offset


def main():
    if not user32.OpenClipboard(0):
        print("ERROR: Could not open clipboard")
        return

    try:
        # Find Chromium Web Custom MIME Data Format
        fmt = 0
        target_fmt = None
        while True:
            fmt = user32.EnumClipboardFormats(fmt)
            if fmt == 0:
                break
            buf = ctypes.create_unicode_buffer(256)
            if user32.GetClipboardFormatNameW(fmt, buf, 256):
                if buf.value == "Chromium Web Custom MIME Data Format":
                    target_fmt = fmt
                    break

        if target_fmt is None:
            print("No 'Chromium Web Custom MIME Data Format' found on clipboard.")
            print("Make sure you copied from Slack first.")
            return

        print("=" * 70)
        print("CHROMIUM WEB CUSTOM MIME DATA FORMAT (id={})".format(target_fmt))
        print("=" * 70)

        data = get_clipboard_data(target_fmt)
        if data is None:
            print("Could not read clipboard data.")
            return

        print("  Total size: {} bytes".format(len(data)))
        decode_chromium_pickle(data)

    finally:
        user32.CloseClipboard()


if __name__ == '__main__':
    main()
