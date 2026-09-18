#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/extglob_patterns_after_shopt
# extglob must be enabled before the line containing the pattern is parsed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
[[ abc == +([a-c]) ]] || fail "+( )"
[[ abc == !(xyz) ]] || fail "!( )"
[[ abc == !(abc) ]] && fail "!(abc) must not match abc"
[[ ac == a?(b)c ]] || fail "?( ) zero occurrences"
[[ abbc == a*(b)c ]] || fail "*( ) many occurrences"
[[ ac == @(a|b)c ]] || fail "@( ) alternatives"
echo PASS
exit 0
