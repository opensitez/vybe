#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_invalid_formats/function
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if printf "%${2}" "2" >/dev/null 2>&1; then
  fail "invalid format unexpectedly succeeded"
fi
if ! printf "%s" "ok2" >/dev/null 2>&1; then
  fail "valid format unexpectedly failed"
fi
echo PASS
exit 0
