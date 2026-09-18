#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/match_is_unanchored_substring_search
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ xabcx =~ abc ]] || fail "substring must match"
[[ xabcx =~ ^abc$ ]] && fail "anchors restrict to the whole string"
[[ abc =~ ^abc$ ]] || fail "anchored exact match"
echo PASS
exit 0
