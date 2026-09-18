#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_does_not_mutate_underlying_variable
# Alternate value expansions produce no assignment side effects on the underlying variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="original_state"
res="${x:+new_alternate}"
[ "$res" = "new_alternate" ] || fail "expansion mismatch"
[ "$x" = "original_state" ] || fail "underlying variable was modified: got [$x]"
echo PASS
exit 0
