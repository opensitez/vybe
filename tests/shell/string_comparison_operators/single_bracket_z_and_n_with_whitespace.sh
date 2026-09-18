#!/usr/bin/env bash
# vybe-test: bash/string_comparison_operators/single_bracket_z_and_n_with_whitespace
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
e=; x=val; unset u
[ -z "$e" ] && [ -z "$u" ] && [ -n "$x" ] || fail "basic -z/-n in [ ]"
[ -z " " ] && fail "a single space is not empty"
[ -n " " ] || fail "a single space is non-empty"
[ -n "$e" ] && fail "-n on empty must be false"
echo PASS
exit 0
