#!/bin/zsh
set -u

BASE_DIR="$(cd "$(dirname "$0")" && pwd)"
ZSH_SCRIPT="${BASE_DIR}/start_app.zsh"

if [[ ! -f "${ZSH_SCRIPT}" ]]; then
  echo "Startup failed: start_app.zsh not found: ${ZSH_SCRIPT}"
  echo "Press Enter to exit..."
  read -r _
  exit 1
fi

if [[ ! -x "${ZSH_SCRIPT}" ]]; then
  chmod +x "${ZSH_SCRIPT}" 2>/dev/null || true
fi

/bin/zsh "${ZSH_SCRIPT}"
EXIT_CODE=$?

if [[ ${EXIT_CODE} -ne 0 ]]; then
  echo ""
  echo "Startup failed with exit code ${EXIT_CODE}."
  echo "Press Enter to exit..."
  read -r _
fi

exit ${EXIT_CODE}
