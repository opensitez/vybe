#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/assignment_in_condition_persists
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if v=$(echo captured); then :; fi
[ "$v" = captured ] || fail "got [$v]"
if (( (w = 5) == 0 )); then :; fi
[ "$w" -eq 5 ] || fail "assignment inside a false arithmetic condition still happens, w=$w"
echo PASS
exit 0
