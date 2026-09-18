#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/elif_conditions_stop_at_first_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
n=0
probe() { n=$((n+1)); [ "$1" = t ]; }
if probe f; then r=1; elif probe t; then r=2; elif probe t; then r=3; else r=4; fi
[ "$r" -eq 2 ] || fail "branch $r"
[ "$n" -eq 2 ] || fail "conditions evaluated: want 2 got $n"
echo PASS
exit 0
