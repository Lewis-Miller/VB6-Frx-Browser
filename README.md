# VB6 Frx Browser

A pure VB6 utility for browsing, previewing, and exporting binary resources from Visual Basic 6 form files (`.frm` / `.frx`), projects (`.vbp`), UserControls, Property Pages, and ActiveX Designers.

It parses the various binary formats VB6 uses to store pictures, icons, ImageLists, PropertyBags, strings, lists, and custom ActiveX data, then lets you view and export them cleanly.

## Features

- Open `.vbp`, `.frm`, `.frx`, `.ctl`, `.pag`, `.dsr`, `.ctx`, `.pgx`, `.dsx` files
- Automatically resolves companion files (e.g. load a `.frx` and it finds the matching form)
- Extracts and previews:
  - Pictures (BMP, PNG, JPG, GIF, WMF, EMF, ICO, CUR)
  - Multi-resolution icons (unpacks or packs back into a single `.ico`)
  - ImageList / Toolbar strips (export as individual files or horizontal PNG strip)
  - Strings, multi-line text, RTF, Combo/List `.List` and `.ItemData`
  - VB6 PropertyBags (decoded to readable key=value text)
  - Custom ActiveX blobs, including RIIFS, WAVS, MIDI and OLE
- Hex viewer for any resource
- Batch export with options:
  - Pack multi-res icons into one `.ico`
  - Export ImageLists as PNG strips
  - Group by form/designer folder
  - Skip unindexed raw blobs
- Standalone FRX carving (finds pictures even without the matching `.frm`)
- GDI+ based image loading and strip generation

## Project Structure

| File | Purpose |
|------|---------|
| `frmMain.frm` | Main browser UI |
| `frmViewer.frm` | Hex dump viewer |
| `frmExportOpts.frm` | Export options dialog |
| `modVars.bas` | Core FRX parsing, extraction, and resource types |
| `modGDIPlusLite.bas` | Lightweight GDI+ helpers for preview and strip export |
| `modCommDlg.bas` | Open/Save/Browse folder dialogs |
| `modFileAll.bas` | File path and helper functions |
| `modMain.bas` | Startup and settings helpers |
| `ProjectFRXViewer.vbp` | Project file |

## Building

Open `ProjectFRXViewer.vbp` in the VB6 IDE and compile. No external OCX dependencies are required for the core functionality (uses pure API for dialogs and GDI+).

## Notes

- Control array index tracking uses a simple stack so nested containers and `Index` properties are handled correctly.
- Picture headers cover the common VB6 variants (standard 12-byte, CLSID_StdPicture, OLE streams used by AlphaImageControl, etc.).
- PropertyBag decoding and icon pack/unpack are pure VB6.

## Author

Lewis Miller (Fire Hades)
