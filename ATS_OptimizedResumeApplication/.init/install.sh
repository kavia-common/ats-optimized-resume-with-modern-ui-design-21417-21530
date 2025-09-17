#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
# evidence
command -v node >/dev/null 2>&1 && node -v || true
command -v npm >/dev/null 2>&1 && npm -v || true
# optionally install office-addin-cli: prefer npx/local unless OFFICE_ADDIN_GLOBAL=1
if [ "${INSTALL_OFFICEJS:-0}" = "1" ]; then
  CLI_VER="${OFFICE_ADDIN_CLI_VERSION:-latest}"
  if [ "${OFFICE_ADDIN_GLOBAL:-0}" = "1" ]; then
    # check existing version and skip if matches/pinned
    if command -v office-addin >/dev/null 2>&1 || command -v office-addin-cli >/dev/null 2>&1; then
      echo 'office-addin present; skipping global install'
    else
      sudo npm i -g --no-audit --silent "office-addin-cli@${CLI_VER}" || { echo 'ERROR: npm -g install office-addin-cli failed' >&2; exit 40; }
      NPM_G_BIN=$(npm bin -g 2>/dev/null || true)
      if [ -n "$NPM_G_BIN" ]; then
        PROFILE_FILE=/etc/profile.d/officejs_path.sh
        ENTRY="export PATH=\"${NPM_G_BIN}:\$PATH\""
        # ensure we can write the profile file atomically
        if sudo test -w /etc/profile.d/ 2>/dev/null || sudo test -f "$PROFILE_FILE" 2>/dev/null; then
          echo "$ENTRY" | sudo tee "$PROFILE_FILE" >/dev/null && sudo chmod 644 "$PROFILE_FILE" || { echo 'ERROR: failed to write /etc/profile.d entry' >&2; exit 41; }
        else
          # fallback to workspace-local wrapper
          mkdir -p "$WS_PATH/.env.d" && echo "$ENTRY" > "$WS_PATH/.env.d/officejs_path.sh"
        fi
      fi
    fi
  else
    # prefer npx or local install: create helper script to use npx
    mkdir -p "$WS_PATH/.env.d"
    echo "# helper to run office-addin-cli via npx (no global install)" > "$WS_PATH/.env.d/officejs_npx.sh"
    echo "run_office_addin_cli() { npx --yes --package office-addin-cli${CLI_VER:+@${CLI_VER}} office-addin \"\$@\"; }" >> "$WS_PATH/.env.d/officejs_npx.sh"
  fi
fi
# generate packaging script that computes WS at runtime (avoids injected WS_PATH bug)
cat > scripts/package_xlsm.sh <<'PKG'
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
PKG
# syntax check generated script
bash -n scripts/package_xlsm.sh || { echo 'ERROR: package_xlsm.sh syntax invalid' >&2; exit 46; }
chmod +x scripts/package_xlsm.sh
