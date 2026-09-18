#!/usr/bin/env bash
# vybe-test: bash/double_bracket_conditionals/unary_operator_without_operand_is_syntax_error
# Unlike [ -n ], which is a one-argument string test, [[ -n ]] does not parse.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval '[[ -n ]]' 2>&1 ); st=$?
[ "$st" -ne 0 ] || fail "must fail"
[[ $msg == *"unexpected argument"* ]] || fail "got [$msg]"
echo PASS
exit 0
