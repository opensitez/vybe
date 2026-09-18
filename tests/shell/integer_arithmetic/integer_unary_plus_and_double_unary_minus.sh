#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_unary_plus_and_double_unary_minus
# Unary + is a no-op; unary -- is parsed as plus/minus application.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((+7)) -eq 7 ] || fail "+7 got $((+7))"
[ $((- -7)) -eq 7 ] || fail "- -7 got $((- -7))"
echo PASS
exit 0
