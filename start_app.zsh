#!/bin/zsh
set -u

BASE_DIR="$(cd "$(dirname "$0")" && pwd)"
BRIDGE_HOST="127.0.0.1"
BRIDGE_PORT=8765
BRIDGE_PID=""

cleanup() {
  if [[ -n "${BRIDGE_PID}" ]]; then
    if kill -0 "${BRIDGE_PID}" 2>/dev/null; then
      kill "${BRIDGE_PID}" 2>/dev/null || true
      sleep 0.2
      if kill -0 "${BRIDGE_PID}" 2>/dev/null; then
        kill -9 "${BRIDGE_PID}" 2>/dev/null || true
      fi
    fi
  fi
}

on_exit() {
  cleanup
}
trap on_exit EXIT INT TERM

find_python() {
  if command -v python3 >/dev/null 2>&1; then
    echo "python3"
    return 0
  fi
  if command -v python >/dev/null 2>&1; then
    echo "python"
    return 0
  fi
  return 1
}

wait_port_open() {
  local host="$1"
  local port="$2"
  local i
  for i in {1..60}; do
    if nc -z "${host}" "${port}" >/dev/null 2>&1; then
      return 0
    fi
    sleep 0.2
  done
  return 1
}

main() {
  local py_cmd
  py_cmd="$(find_python)" || {
    echo "Startup failed: python3/python not found in PATH."
    return 1
  }

  local bridge_script="${BASE_DIR}/src/bridge_server.py"
  if [[ ! -f "${bridge_script}" ]]; then
    echo "Startup failed: Bridge script not found: ${bridge_script}"
    return 1
  fi

  cd "${BASE_DIR}" || return 1
  "${py_cmd}" "${bridge_script}" >/dev/null 2>&1 &
  BRIDGE_PID=$!

  wait_port_open "${BRIDGE_HOST}" "${BRIDGE_PORT}" || {
    echo "Startup failed: Bridge server did not start on port ${BRIDGE_PORT}."
    return 1
  }

  local index_path="${BASE_DIR}/index.html"
  if [[ ! -f "${index_path}" ]]; then
    echo "Startup failed: index.html not found: ${index_path}"
    return 1
  fi

  open "${index_path}" >/dev/null 2>&1 || {
    echo "Could not open browser automatically. Please open ${index_path} manually."
  }

  echo "Opened in default browser: ${index_path}"
  echo "Bridge server running: http://${BRIDGE_HOST}:${BRIDGE_PORT}"
  echo "Press Enter here to stop the bridge server after you finish."
  read -r _
  return 0
}

main
