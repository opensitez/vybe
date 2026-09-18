#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/empty_pattern_matches_only_the_empty_string
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
e=
[[ "" == $e ]] || fail "empty matches empty"
[[ a == $e ]] && fail "empty pattern must not match a"
echo PASS
exit 0
