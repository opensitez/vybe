#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_lazy_evaluation_skips_arithmetic_mutation
# If the variable is set and non-null, arithmetic mutations in default word are never evaluated.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="exists"
counter=10
res="${x:-$(( counter += 5 ))}"
[ "$res" = "exists" ] || fail "expansion mismatch: got [$res]"
[ "$counter" -eq 10 ] || fail "arithmetic in default was evaluated eagerly: got $counter"
echo PASS
exit 0
