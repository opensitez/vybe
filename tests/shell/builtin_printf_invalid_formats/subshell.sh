#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_invalid_formats/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if printf "%${3}" "3" >/dev/null 2>&1; then
  fail "invalid format unexpectedly succeeded"
fi
if ! printf "%s" "ok3" >/dev/null 2>&1; then
  fail "valid format unexpectedly failed"
fi
echo PASS
exit 0
