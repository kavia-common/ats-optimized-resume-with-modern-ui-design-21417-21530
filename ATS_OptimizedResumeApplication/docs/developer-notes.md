# Developer Notes: ATS Optimized Resume Application (Excel VBA)

This codebase provides a maintainable, modular VBA implementation for an interactive resume builder optimized for ATS parsing. Use a Windows host with Excel to import the modules into a .xlsm workbook.

Key steps:
- Import all vba/*.bas modules.
- In ThisWorkbook:
  Private Sub Workbook_Open(): EntryPoint.App_Initialize
  Private Sub Workbook_BeforeClose(Cancel As Boolean): EntryPoint.App_Shutdown
- Create a visible "UI" sheet. Add Form Controls (buttons) and assign macros from UI.bas and EntryPoint.bas.
- Internal sheets are created automatically: Data, VersionHistory, HelpDocs, Templates (very hidden).

Feature mapping:
- Guided forms: UI.SetupUI + UI.UI_SectionChanged + UI.UI_SaveSection
- ATS optimization: ATSEngine.* + UI.UpdateKeywordInsights
- Formatting enforcement: Templates.* + ATSEngine.EnforceFormattingRules
- Version control: VersionControl.* (auto-save, manual save, history)
- Accessibility: Accessibility.* (Undo/Redo, protection)
- Export/Print: ExportPrint.* (BuildExportSheet, ExportActiveVersionToPDF)
- Documentation: Documentation.BuildDocumentation + Templates.ExportStyleGuide

CI constraints:
- Linux CI package script produces a placeholder .xlsm without a valid vbaProject.bin.
- Test on Windows Excel for full macro execution.

Extensibility:
- Add new templates in Templates.EnsureTemplates.
- Extend keyword banks in ATSEngine.BuildKeywordBank (or load from HelpDocs).
- Add new sections by updating the DataModel section order key and wiring UI dropdown.
