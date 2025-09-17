#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
mkdir -p "$WS_PATH" && cd "$WS_PATH"
mkdir -p templates vba scripts tests ci docs
cat > vba/sample_module.bas <<'VBA'
Attribute VB_Name = "SampleModule"
Sub Hello()
    ' sample macro
End Sub
VBA
export WS_PATH="$WS_PATH"
python3 - <<'PY'
import os
from openpyxl import Workbook
ws_path = os.environ['WS_PATH']
wb = Workbook()
ws = wb.active
ws.title = 'Resume'
ws['A1'] = 'Name'
out = os.path.join(ws_path, 'templates', 'sample_template.xlsx')
wb.save(out)
print(out)
PY
cat > README.md <<'RD'
Project workspace for ATS_OptimizedResumeApplication.
Templates: templates/ (XLSX/XLSM placeholders); VBA modules: vba/; packaging scripts: scripts/.
RD
cat > docs/ci.md <<'RD'
Excel must run on a Windows host. The container provides file-based packaging/validation only.
Do NOT install pywin32 inside this Linux container. To use pywin32, run tests on a Windows host and set INSTALL_PYWIN32=1 there.
To embed a real vbaProject.bin: open a valid .xlsm on Windows, extract vbaProject.bin and place it in templates/ replacing the placeholder.
Environment variables:
  - INSTALL_OFFICEJS=1 to install office-addin-cli tooling (optional)
  - OFFICE_ADDIN_CLI_VERSION to pin office-addin-cli version (optional)
  - OFFICE_ADDIN_GLOBAL=1 to perform a global npm -g install (optional; default avoids global install)
  - INSTALL_PYWIN32=1 to install pywin32 (Windows host only)
  - UPGRADE_PIP=1 to upgrade pip (not recommended by default)
  - CONTAINER_WORKSPACE to override default workspace path
RD
