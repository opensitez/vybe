#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/sign_applies_to_base_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((-0x10)) -eq -16 ] || fail "got $((-0x10))"
[ $((-2#11)) -eq -3 ] || fail "got $((-2#11))"
[ $((-010)) -eq -8 ] || fail "got $((-010))"
echo PASS
exit 0
