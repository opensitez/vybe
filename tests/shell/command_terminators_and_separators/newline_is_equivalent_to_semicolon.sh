#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/newline_is_equivalent_to_semicolon
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=$(echo 1; echo 2)
b=$(echo 1
echo 2)
[ "$a" = "$b" ] || fail "[$a] vs [$b]"
echo PASS
exit 0
