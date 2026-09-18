#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/bash_rematch_holds_whole_match_and_groups
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc123
[[ $x =~ ^([a-z]+)([0-9]+)$ ]] || fail "no match"
[ "${BASH_REMATCH[0]}" = abc123 ] || fail "[0] got [${BASH_REMATCH[0]}]"
[ "${BASH_REMATCH[1]}" = abc ] || fail "[1] got [${BASH_REMATCH[1]}]"
[ "${BASH_REMATCH[2]}" = 123 ] || fail "[2] got [${BASH_REMATCH[2]}]"
[ "${#BASH_REMATCH[@]}" -eq 3 ] || fail "count got ${#BASH_REMATCH[@]}"
echo PASS
exit 0
