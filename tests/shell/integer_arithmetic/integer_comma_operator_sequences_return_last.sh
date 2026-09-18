#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_comma_operator_sequences_return_last
# In comma expressions, statements run left-to-right and the last value is returned.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
result=$((x = 7, x + 2, x * 2))
[ "$result" -eq 16 ] || fail "result want 16, got $result"
[ "$x" -eq 7 ] || fail "x should remain 7, got $x"
echo PASS
exit 0
