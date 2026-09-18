#!/usr/bin/env bash
# feature: exam-variants
#
# Run the test suite.
#
#   ./run_tests.sh                          all 58 tests, quietly
#   ./run_tests.sh -v                       one line per test
#   ./run_tests.sh test_testmess.FunctionalTest        one class
#   ./run_tests.sh -v test_testmess.FunctionalTest.test_end_to_end_documents
#
# Arguments are handed straight to unittest; the test module is added for you
# when you do not name one.  Set PYTHON=... to pick a specific interpreter.
set -euo pipefail

cd "$(dirname "$0")"

MODULE=test_testmess

if [ -z "${PYTHON:-}" ]; then
  if command -v python3 >/dev/null 2>&1; then
    PYTHON=python3
  elif command -v python >/dev/null 2>&1; then
    PYTHON=python
  else
    echo "error: no python found on PATH; install Python 3.8+ (see README)" >&2
    exit 1
  fi
fi

# Anything that is not a flag is a test to run, so only add the module when
# the caller named none -- that way "./run_tests.sh -v" still works.
args=("$@")
for arg in "$@"; do
  case "$arg" in
    -*) ;;
    *) MODULE=''; break ;;
  esac
done
if [ -n "$MODULE" ]; then
  args+=("$MODULE")
fi

echo "== $("$PYTHON" --version 2>&1) at $(command -v "$PYTHON")"
exec "$PYTHON" -m unittest "${args[@]}"
