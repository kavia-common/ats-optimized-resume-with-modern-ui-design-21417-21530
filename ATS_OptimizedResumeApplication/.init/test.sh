#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
PORT="${VALIDATION_PORT:-8008}"
TIMEOUT="${VALIDATION_TIMEOUT:-10}"
SLEEP=0.5
TARGET_PATH="/sample_template_with_vba.xlsm"
HTTP_CLIENT=none
if command -v curl >/dev/null 2>&1; then HTTP_CLIENT=curl; elif command -v wget >/dev/null 2>&1; then HTTP_CLIENT=wget; else echo 'ERROR: neither curl nor wget available' >&2; exit 61; fi
COUNT=0
ITER=$(( (TIMEOUT*1000) / (SLEEP*1000) ))
if [ $ITER -lt 1 ]; then ITER=1; fi
HTTP_STATUS=0
while [ $COUNT -lt $((ITER + 1)) ]; do
  if [ "$HTTP_CLIENT" = "curl" ]; then
    HTTP_STATUS=$(curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:${PORT}${TARGET_PATH} || true)
  else
    HTTP_STATUS=$(wget -q --spider --server-response http://127.0.0.1:${PORT}${TARGET_PATH} 2>&1 | awk '/HTTP\//{print $2; exit}' || true)
  fi
  if [ "$HTTP_STATUS" = "200" ]; then
    echo "VALIDATION_PASSED: http status=$HTTP_STATUS" | tee templates/validation.out
    # append a bit of server log for evidence
    tail -n 200 templates/server.log >> templates/validation.out || true
    exit 0
  fi
  sleep $SLEEP
  COUNT=$((COUNT+1))
done
# failed
echo "ERROR: http serve failed status=$HTTP_STATUS; see templates/server.log" | tee templates/validation.out >&2
exit 62
