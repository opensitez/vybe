#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/bare_name_reads_the_variable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=5
[ $((x + 1)) -eq 6 ] || fail "got $((x + 1))"
[ $((x * x)) -eq 25 ] || fail "got $((x * x))"
echo PASS
exit 0
