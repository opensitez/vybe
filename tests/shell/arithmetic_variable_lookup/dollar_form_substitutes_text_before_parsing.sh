#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/dollar_form_substitutes_text_before_parsing
# $x pastes "1+2" into the expression (1+2*2 = 5); a bare x evaluates the
# variable first ((1+2)*2 = 6).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='1+2'
[ $(( $x * 2 )) -eq 5 ] || fail "\$x form want 5 got $(( $x * 2 ))"
[ $(( x * 2 )) -eq 6 ] || fail "bare form want 6 got $(( x * 2 ))"
echo PASS
exit 0
