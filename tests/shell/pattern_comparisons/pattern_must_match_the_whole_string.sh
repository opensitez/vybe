#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/pattern_must_match_the_whole_string
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ abc == b ]] && fail "b is not the whole string"
[[ abc == *b* ]] || fail "*b* covers the whole string"
[[ abc == a?c ]] || fail "a?c"
[[ abc == a?? ]] || fail "a??"
[[ abc == a? ]] && fail "a? is too short"
echo PASS
exit 0
