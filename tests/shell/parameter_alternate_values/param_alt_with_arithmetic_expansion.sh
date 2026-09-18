#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_with_arithmetic_expansion
# The alternate word evaluates inline arithmetic expressions $(( ... )) when condition is met.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
compute_flag="yes"
res="${compute_flag:+$(( 100 * 2 + 50 ))}"
[ "$res" -eq 250 ] || fail "arithmetic evaluation in alternate failed: got [$res]"
echo PASS
exit 0
