#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_logical_negation_exclamation
# The '!' operator inverts the truth value of a conditional expression inside [[ ... ]].
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ! 1 -eq 2 ]] || fail "! (1 -eq 2) should be true"
[[ ! 1 -eq 1 ]] && fail "! (1 -eq 1) should be false"
[[ ! ( "a" == "b" || 2 -eq 3 ) ]] || fail "! (false || false) should be true"
echo PASS
exit 0
