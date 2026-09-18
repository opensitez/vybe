#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_continues_after_logical_or
# A newline immediately following '||' continues the conditional list to the next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
flag="unset"
false ||
    flag="fallback_reached"
[ "$flag" = "fallback_reached" ] || fail "newline after || failed: got [$flag]"
echo PASS
exit 0
