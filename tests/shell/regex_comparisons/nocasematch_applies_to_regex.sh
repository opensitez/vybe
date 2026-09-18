#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/nocasematch_applies_to_regex
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ABC =~ ^abc$ ]] && fail "default case sensitive"
shopt -s nocasematch
[[ ABC =~ ^abc$ ]] || fail "nocasematch regex"
echo PASS
exit 0
