#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/leading_separator_is_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '; echo a' 2>/dev/null; a=$?
eval '&& echo a' 2>/dev/null; b=$?
eval '| echo a' 2>/dev/null; c=$?
[ "$a" -ne 0 ] || fail "leading ; must fail"
[ "$b" -ne 0 ] || fail "leading && must fail"
[ "$c" -ne 0 ] || fail "leading | must fail"
echo PASS
exit 0
