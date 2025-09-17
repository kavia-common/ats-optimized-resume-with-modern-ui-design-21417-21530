#!/usr/bin/env bash
set -euo pipefail
# compute workspace: parent of scripts/ (assumes scripts/ lives under workspace)
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WS="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$WS"
TMP_DIR=$(mktemp -d)
trap 'rc=$?; rm -rf "${TMP_DIR}" || true; exit ${rc}' EXIT
cp "templates/sample_template.xlsx" "${TMP_DIR}/template.xlsx" || { echo 'ERROR: templates/sample_template.xlsx not found' >&2; exit 42; }
cd "$TMP_DIR"
unzip -q template.xlsx -d unpacked || { echo 'ERROR: unzip template failed' >&2; exit 43; }
mkdir -p unpacked/xl
printf 'DUMMY VBAPROJ - placeholder; not a valid VBA binary' > unpacked/xl/vbaProject.bin
cat > unpacked/NOTICE.txt <<'NOTE'
This .xlsm was generated inside a headless Linux container as a CI placeholder. It contains a dummy vbaProject.bin and may be rejected by Excel. To embed a valid VBA project, produce vbaProject.bin from a real .xlsm on Windows and replace this file.
NOTE
cd unpacked
zip -qr "${WS}/templates/sample_template_with_vba.xlsm" . || { echo 'ERROR: zip creation failed' >&2; exit 44; }
# explicit validation: ensure produced file is a zip
unzip -l "${WS}/templates/sample_template_with_vba.xlsm" >/dev/null 2>&1 || { echo 'ERROR: produced xlsm is not a valid zip' >&2; exit 45; }
