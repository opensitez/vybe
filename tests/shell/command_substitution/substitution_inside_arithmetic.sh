#!/usr/bin/env bash
# vybe-test: bash/command_substitution/substitution_inside_arithmetic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
n=$(( $(echo 2) + 1 ))
[ "$n" -eq 3 ] || fail "want 3 got $n"
m=$(( $(echo 6) / $(echo 2) ))
[ "$m" -eq 3 ] || fail "want 3 got $m"
echo PASS
exit 0
