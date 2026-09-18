#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/double_bracket_empty_operand_is_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset u; e=
[[ "" -eq 0 ]] || fail "literal empty"
[[ $u -eq 0 ]] || fail "unset variable"
[[ $e -lt 1 ]] || fail "empty variable"
echo PASS
exit 0
