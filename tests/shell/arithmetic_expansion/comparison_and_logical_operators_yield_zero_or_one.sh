#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/comparison_and_logical_operators_yield_zero_or_one
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="$((1<2)) $((3==4)) $((3!=4)) $((2>=2)) $((!0)) $((!5)) $((2&&0)) $((0||3)) $((7&&8))"
[ "$out" = "1 0 1 1 1 0 0 1 1" ] || fail "got [$out]"
echo PASS
exit 0
