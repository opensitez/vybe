#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_invalid_formats/cleanup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if printf "%${15}" "15" >/dev/null 2>&1; then
  fail "invalid format unexpectedly succeeded"
fi
if ! printf "%s" "ok15" >/dev/null 2>&1; then
  fail "valid format unexpectedly failed"
fi
echo PASS
exit 0
