#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_escape_processing/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 3 % 2 == 0 )); then
  out=$(printf "%b" "x\\n")
  [ "$out" = "x\n" ] || fail "escaped newline mismatch"
else
  out=$(printf "%b" "x\\\\n")
  [ "$out" = "x\\n" ] || fail "literal backslash mismatch"
fi
echo PASS
exit 0
