# 📑 CHANGELOG
All notable changes to the WAT Data Automation Tool are documented here.

## v1.0.0 – Initial Release
- First release of the WAT Data Automation Tool.
- Automated .wat → Excel conversion with structured workbook output.
- Generated Per Unit Data, Per Wafer Summary, and Summary Sheet.
- Capability plotting via Histogram Viewer with Cp/Cpk statistics.
- Tkinter GUI for file selection, automation control, and status logging.
➡️ [View v1.0.0 Release](https://github.com/roannelafuente/WAT-Data-Automation)

## v1.1.1 – Multi‑Class Architecture & Dashboard Update
Note: This release skips v1.1.0 because of a packaging error. The next valid version is v1.1.1.
- Transitioned to a modular, multi‑class architecture for cleaner structure and maintainability.
- Integrated centralized logging system:
- Auto‑creates timestamped log files in a dedicated /logs folder.
- Automatically removes logs older than 30 days.
- Unified error capture across all modules with GUI + file logging.
- Updated Histogram Plot Dashboard:
- Cleaner layout and improved parameter navigation.
- Enhanced Cp/Cpk statistics panel.
- Improved ±3σ normal curve overlay.
- GUI refinements:
- Polished spacing and alignment.
- Clearer success/error messages.
- More consistent user feedback.
- Internal improvements:
- Better separation of concerns across parser, builder, summary generator, histogram viewer, and GUI controller.
- More reliable workbook handling and data flow.
➡️ [View v1.1.1 Release](https://github.com/roannelafuente/WAT-Data-Automation-v1.1.1)
