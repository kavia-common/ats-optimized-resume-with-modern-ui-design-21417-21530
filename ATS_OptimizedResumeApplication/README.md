Project workspace for ATS_OptimizedResumeApplication.
Templates: templates/ (XLSX/XLSM placeholders); VBA modules: vba/; packaging scripts: scripts/.

How to use the VBA code in Excel:
1) Open templates/sample_template.xlsx in Excel (Windows recommended).
2) Enable Developer tab -> Visual Basic (VBA Editor).
3) Import all .bas files from the vba/ folder into the workbook (File -> Import File…).
4) In ThisWorkbook code module, add:
   Private Sub Workbook_Open()
       EntryPoint.App_Initialize
   End Sub
   Private Sub Workbook_BeforeClose(Cancel As Boolean)
       EntryPoint.App_Shutdown
   End Sub
5) Create a visible sheet named "UI". Add buttons (Form Controls) and assign these macros:
   - Setup UI: UI.SetupUI
   - Section Changed (assign to dropdown change or call manually): UI.UI_SectionChanged
   - Save Section: UI.UI_SaveSection
   - Add Bullet Prefix: UI.UI_AddBulletPrefix
   - Save Version: UI.UI_NewVersion
   - Compare Versions: UI.UI_CompareVersions
   - Create Branch: UI.UI_CreateBranch
   - Merge Into Active: UI.UI_MergeIntoActive
   - Apply Template: UI.UI_ApplyTemplate
   - Export PDF: UI.UI_ExportPDF
   - Show Help: EntryPoint.App_ShowHelp
6) Save as a macro-enabled workbook (.xlsm). Optionally use scripts/package_xlsm.sh to produce a CI placeholder.
7) Use the UI sheet to manage resume sections, validate content, optimize keywords, and export PDFs.

Module overview:
- Globals.bas: constants, types, helpers, styles, sheet creation, accessibility announcements
- DataModel.bas: hidden Data sheet key-value store, section CRUD, section order
- ATSEngine.bas: keyword bank, scoring, gaps, formatting enforcement, ATS score simulation
- Validation.bas: field validators and contextual help tips
- Templates.bas: template library, style enforcement, style guide export
- VersionControl.bas: auto-save, manual save, history, compare, restore, branching, merge
- ExportPrint.bas: build export sheet and export to PDF
- UI.bas: UI orchestration for guided forms, validation, and optimization insights
- Accessibility.bas: undo/redo and sheet protection
- EntryPoint.bas: initialize and shutdown hooks
- Documentation.bas: HelpDocs content generator

Note:
- This repository cannot embed a real vbaProject.bin on Linux CI. Build and test macros on Windows Excel for full functionality.
