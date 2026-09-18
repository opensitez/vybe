#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/nocasematch_affects_pattern_matching
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ABC == a* ]] && fail "default case sensitive"
shopt -s nocasematch
[[ ABC == a* ]] || fail "nocasematch pattern"
[[ ABC == [a-c]bc ]] || fail "nocasematch range"
echo PASS
exit 0
