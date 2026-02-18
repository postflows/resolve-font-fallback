# Font Fallback

> Part of [PostFlows](https://github.com/postflows) toolkit for DaVinci Resolve

Detect missing fonts on the timeline and replace them with a user-selected font, with restoration tags for reverting.

## What it does

Scans current timeline for Text+ and MultiText fonts, reports missing fonts/styles, batch-replaces with a chosen font/style, embeds restoration tags in comments, and can restore original fonts from those tags. Logs replacements to desktop JSON.

## Requirements

- DaVinci Resolve 18+
- **Optional:** `pyperclip` for copying the “missing fonts” list to the clipboard. Without it, “Copy Missing” is unavailable; all other features work.

## Installation

### 1. Copy the script

Copy the script to Resolve’s Fusion Scripts folder:

- **macOS:** `~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/`
- **Windows:** `C:\ProgramData\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\`

### 2. Optional: pyperclip (for clipboard copy)

To enable **Copy Missing** (copy list of missing fonts to clipboard), install `pyperclip`. If you don’t have Python installed, use **Python 3.11–3.13** from [python.org](https://www.python.org/downloads/) (on Windows, check “Add Python to PATH”). Then try the standard method:

```bash
pip3 install pyperclip
```

Then run the script; if “Copy Missing” works, you’re done. If Resolve reports that pyperclip is missing, it may be using its own Python — then install into that interpreter (see Resolve **Preferences → System → General → External scripting** for the path). Example for Resolve’s bundled Python on macOS:

```bash
/Applications/DaVinci\ Resolve/DaVinci\ Resolve.app/Contents/Libraries/Frameworks/Python.framework/Versions/3.10/bin/python3 -m pip install pyperclip
```

Windows (adjust path if needed):

```cmd
"C:\Program Files\Blackmagic Design\DaVinci Resolve\python.exe" -m pip install pyperclip
```

If you don’t install pyperclip, the script runs normally; only the clipboard copy action will be disabled or show a message.

## Usage

Run from Workspace > Scripts. Choose replacement font/style, click Refresh Timeline, then Replace Missing or Restore Fonts; Copy Missing copies list to clipboard.

## License

MIT © PostFlows
