#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/failed_match_empties_bash_rematch
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ abc =~ (b) ]] || fail "setup match"
[ "${#BASH_REMATCH[@]}" -eq 2 ] || fail "setup count"
[[ abc =~ z ]]; st=$?
[ "$st" -eq 1 ] || fail "non-match status want 1 got $st"
[ "${#BASH_REMATCH[@]}" -eq 0 ] || fail "BASH_REMATCH must be empty, has ${#BASH_REMATCH[@]}"
echo PASS
exit 0
