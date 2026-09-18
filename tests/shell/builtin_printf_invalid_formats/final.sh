#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_invalid_formats/final
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if printf "%${20}" "20" >/dev/null 2>&1; then
  fail "invalid format unexpectedly succeeded"
fi
if ! printf "%s" "ok20" >/dev/null 2>&1; then
  fail "valid format unexpectedly failed"
fi
echo PASS
exit 0
