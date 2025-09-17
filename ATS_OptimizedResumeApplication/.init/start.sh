#!/usr/bin/env bash
set -euo pipefail
WS_PATH="${CONTAINER_WORKSPACE:-/home/kavia/workspace/code-generation/ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication}"
cd "$WS_PATH"
PORT="${VALIDATION_PORT:-8008}"
mkdir -p templates
# start server and record logs and pid
python3 -m http.server "$PORT" --directory templates > templates/server.log 2>&1 &
PID=$!
echo "$PID" > templates/server.pid
trap 'kill ${PID} >/dev/null 2>&1 || true' EXIT INT TERM
sleep 0.05
printf "SERVER_STARTED pid=%s port=%s\n" "$PID" "$PORT" > templates/server.start.info
