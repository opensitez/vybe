#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_integer_attribute_declaration_dash_i
# Declaring 'local -i num' applies the integer attribute locally, evaluating assignments as arithmetic.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
int_fn() {
    local -i num="10 + 20"
    [ "$num" -eq 30 ] || fail "local integer arithmetic failed: got $num"
    num+="5"
    [ "$num" -eq 35 ] || fail "local integer addition failed: got $num"
}
int_fn
echo PASS
exit 0
