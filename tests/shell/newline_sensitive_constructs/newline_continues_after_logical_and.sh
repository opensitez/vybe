#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_continues_after_logical_and
# A newline immediately following '&&' continues the conditional list without requiring a backslash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val=0
true &&
    val=10
[ "$val" -eq 10 ] || fail "newline after && failed: want 10, got $val"
echo PASS
exit 0
