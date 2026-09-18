#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/empty_unset_and_null_operands_are_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset u; e=
[ "$(( ))" -eq 0 ] || fail "empty expression got $(( ))"
[ "$((u))" -eq 0 ] || fail "unset got $((u))"
[ "$((e))" -eq 0 ] || fail "null got $((e))"
[ "$((u + 4))" -eq 4 ] || fail "unset in sum got $((u + 4))"
echo PASS
exit 0
