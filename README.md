# Wol

VB6 Mini Discover (`MiniDisc.exe` / `Minidisc.vbp`) that loads `audit.xml` (WMI classes to query), runs discovery against a computer via WMI/ADSI/MSXML, and writes discovery output (`discover.xml` / related assets). Folder name is historical `Wol`. Open `Minidisc.vbp` in the VB6 IDE.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `MiniDisc` (`Minidisc.vbp`) | VB6 | WinForms exe | Load audit.xml, query WMI classes, write discovery output |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `Minidisc.vbp`

## Requirements

- Visual Basic 6.0 IDE
- WMI / Active DS / MSXML as referenced by the project

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/Wol`.
Company names in `.vbp` files: Dave Robinson. Portions noted from NAT in version info.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
