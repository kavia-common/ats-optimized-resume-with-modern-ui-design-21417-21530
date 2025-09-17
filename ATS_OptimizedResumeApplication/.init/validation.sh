#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
# run build, start, test, stop in sequence; ensure stop runs on exit
trap 'bash "$WS_PATH/scripts/validation_stop.sh" >/dev/null 2>&1 || true' EXIT INT TERM
# prefer using the canonical script names in workspace/scripts if present, otherwise use the local ones
# copy the minimal scripts into workspace/scripts for canonical naming consistency
mkdir -p scripts
cat > scripts/validation_build.sh <<'B'
#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
bash scripts/package_xlsm.sh
[ -f templates/sample_template_with_vba.xlsm ] || { echo 'ERROR: build failed (missing xlsm)' >&2; exit 60; }
B
cat > scripts/validation_start.sh <<'S'
#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
PORT="${VALIDATION_PORT:-8008}"
mkdir -p templates
python3 -m http.server "$PORT" --directory templates > templates/server.log 2>&1 &
PID=$!
echo "$PID" > templates/server.pid
S
cat > scripts/validation_test.sh <<'T'
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
    tail -n 200 templates/server.log >> templates/validation.out || true
    exit 0
  fi
  sleep $SLEEP
  COUNT=$((COUNT+1))
done
echo "ERROR: http serve failed status=$HTTP_STATUS; see templates/server.log" | tee templates/validation.out >&2
exit 62
T
cat > scripts/validation_stop.sh <<'P'
#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
PID_FILE=templates/server.pid
if [ -f "$PID_FILE" ]; then
  PID=$(cat "$PID_FILE" || true)
  if [ -n "$PID" ]; then kill "$PID" >/dev/null 2>&1 || true; sleep 0.1; fi
  rm -f "$PID_FILE" || true
fi
P
chmod +x scripts/validation_*.sh
# run sequence
scripts/validation_build.sh
scripts/validation_start.sh
# give server a short moment
sleep 0.1
if ! scripts/validation_test.sh; then
  echo 'VALIDATION_FAILED' >&2
  scripts/validation_stop.sh || true
  exit 62
fi
scripts/validation_stop.sh
echo 'VALIDATION_COMPLETE'
