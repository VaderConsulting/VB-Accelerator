# VB Accelerator

Steve McMahon vbAccelerator VB6 controls and samples: S-Grid, Image List, List Bar, popup menu DLL, SysTray, journal record hook, icon extractor, transparent menu, and multi-monitor helpers.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** ActiveX DLL, ActiveX OCX, WinForms exe

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `vbAcceleratorGrid6` (`S Grid Control/pVBALGrid.vbp`) | VB6 | ActiveX OCX | vbAccelerator S-Grid Control |
| `pTestSGrid` (`S Grid Control/pTest.vbp`) | VB6 | WinForms exe | vbAccelerator S-Grid Demonstration |
| `pJournalRecordHook` (`VB6_Journal_Record_Hook_Sample/pJournalRecordHook6.vbp`) | VB6 | WinForms exe | Journal Record Hook Demonstration |
| `Project1` (`CRC-32/CRCTest.vbp`) | VB6 | WinForms exe | Project1 |
| `SysTray` (`SysTray_Demonstration_Project/PSYSTRAY.VBP`) | VB6 | WinForms exe | vbAccelerator SysTray Sample |
| `TestImageList` (`vbalilsc6/pTestIml6.vbp`) | VB6 | WinForms exe | vbAccelerator Image List Test Project |
| `vbalIml6` (`vbalilsc6/vbalIml6.vbp`) | VB6 | ActiveX OCX | vbAccelerator Image List Control |
| `TestImageList` (`vbalilsd6/pTestIml6.vbp`) | VB6 | WinForms exe | vbAccelerator Image List Test Project |
| `pTestOutlookBar` (`VB6_List_Bar_Full_Source/pTestLbar6.vbp`) | VB6 | WinForms exe | vbAccelerator ListBar Control Tester |
| `vbalLbar6` (`VB6_List_Bar_Full_Source/vbalLBar6.vbp`) | VB6 | ActiveX OCX | vbAccelerator ListBar Control |
| `pIconExtractor6` (`VB6_Icon_Extractor_Source_Code/pIconEx6.vbp`) | VB6 | WinForms exe | vbAccelerator Icon Explorer |
| `PopupMenu` (`cnewmnuc/PopupMenu.vbp`) | VB6 | ActiveX DLL | vbAccelerator PopupMenu Active X DLL |
| `pTestNewMenus` (`cnewmnuc/pTest.vbp`) | VB6 | WinForms exe | vbAccelerator PopupMenu Tester |
| `Project1` (`Side_Logo_Demonstration_Project/TestLogo.vbp`) | VB6 | WinForms exe | Project1 |
| `pTransparentMenu6` (`VB6_Transparent_Menu_Demonstration/pTransparentMenu6.vbp`) | VB6 | WinForms exe | Windows 2000 Transparent Menu Tester |
| `TestMonitors` (`Multiple_Monitor_Support_Sample_Code/pTestMonitors.vbp`) | VB6 | WinForms exe | vbAccelerator Multiple Monitor Code Tester |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `S Grid Control/pVBALGrid.vbp`
- `S Grid Control/pTest.vbp`
- `VB6_Journal_Record_Hook_Sample/pJournalRecordHook6.vbp`
- `CRC-32/CRCTest.vbp`
- `SysTray_Demonstration_Project/PSYSTRAY.VBP`
- `vbalilsc6/pTestIml6.vbp`
- `vbalilsc6/vbalIml6.vbp`
- `vbalilsd6/pTestIml6.vbp`
- `VB6_List_Bar_Full_Source/pTestLbar6.vbp`
- `VB6_List_Bar_Full_Source/vbalLBar6.vbp`
- `VB6_Icon_Extractor_Source_Code/pIconEx6.vbp`
- `cnewmnuc/PopupMenu.vbp`
- … and 4 more `.vbp` files in the tree

## Requirements

- Visual Basic 6.0 IDE
- Registered OCX/DLL dependencies referenced by the `.vbp` (may need to be installed separately):
  - `vbalGrid.ocx`
  - `vbalIml6.ocx`

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/VB Accelerator`.
Company names in `.vbp` files: Hard & Software, aaa, vbAccelerator.
Third-party attribution: Steve McMahon / vbAccelerator. See `THIRD_PARTY_NOTICES.md`.

## License

Third-party code remains under its original terms (or none, where none were supplied). See `THIRD_PARTY_NOTICES.md`. Do not treat this tree as VaderConsulting original MIT-licensed work.
