#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
# run packaging script (must be present and valid)
bash scripts/package_xlsm.sh
[ -f templates/sample_template_with_vba.xlsm ] || { echo 'ERROR: build failed (missing xlsm)' >&2; exit 60; }
 echo "BUILD_OK: templates/sample_template_with_vba.xlsm exists"
