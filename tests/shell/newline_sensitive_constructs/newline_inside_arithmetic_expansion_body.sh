#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_inside_arithmetic_expansion_body
# Newlines within $(( ... )) arithmetic expansion body are treated as ordinary whitespace.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val=$((
    10 +
    25 *
    2
))
[ "$val" -eq 60 ] || fail "multiline arithmetic expansion: want 60, got $val"
echo PASS
exit 0
