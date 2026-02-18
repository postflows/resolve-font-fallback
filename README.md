# Font Fallback

> Part of [PostFlows](https://github.com/postflows) toolkit for DaVinci Resolve

Detect missing fonts on the timeline and replace them with a user-selected font, with restoration tags for reverting.

## What it does

Scans current timeline for Text+ and MultiText fonts, reports missing fonts/styles, batch-replaces with a chosen font/style, embeds restoration tags in comments, and can restore original fonts from those tags. Logs replacements to desktop JSON.

## Requirements

- DaVinci Resolve 18+
- Optional: pyperclip for clipboard copy

## Installation

Copy the script to:

- **macOS:** `~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/`
- **Windows:** `C:\ProgramData\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\`

## Usage

Run from Workspace > Scripts. Choose replacement font/style, click Refresh Timeline, then Replace Missing or Restore Fonts; Copy Missing copies list to clipboard.

## License

MIT © PostFlows
