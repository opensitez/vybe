#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/single_bracket_operands_must_fit_sixty_four_bits
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ 9223372036854775807 -gt 9223372036854775806 ] || fail "max value compares"
[ -9223372036854775808 -lt 0 ] || fail "min value compares"
[ 9223372036854775808 -gt 0 ] 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "out of range must be an error, got $st"
echo PASS
exit 0
