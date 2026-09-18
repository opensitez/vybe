#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/leftmost_longest_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ aaa =~ a+ ]] && [ "${BASH_REMATCH[0]}" = aaa ] || fail "a+ got [${BASH_REMATCH[0]}]"
[[ abcabc =~ b.* ]] && [ "${BASH_REMATCH[0]}" = bcabc ] || fail "b.* got [${BASH_REMATCH[0]}]"
[[ xaay =~ a ]] && [ "${BASH_REMATCH[0]}" = a ] || fail "single a"
echo PASS
exit 0
