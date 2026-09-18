#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/unmatched_optional_group_is_empty_but_counted
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ac =~ ^a(b)?c$ ]] || fail "must match"
[ "${#BASH_REMATCH[@]}" -eq 2 ] || fail "count want 2 got ${#BASH_REMATCH[@]}"
[ -z "${BASH_REMATCH[1]}" ] || fail "group 1 must be empty, got [${BASH_REMATCH[1]}]"
echo PASS
exit 0
