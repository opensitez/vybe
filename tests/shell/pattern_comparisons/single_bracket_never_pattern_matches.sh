#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/single_bracket_never_pattern_matches
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ abc = "a*" ] && fail "[ compares strings literally"
[ 'a*' = "a*" ] || fail "literal equality in ["
echo PASS
exit 0
