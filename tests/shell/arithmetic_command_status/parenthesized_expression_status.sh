#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/parenthesized_expression_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( (1 + 2) * 3 )); st=$?
[ "$st" -eq 0 ] || fail "(( (1 + 2) * 3 )) must succeed"
(( (3 - 3) * (8 + 1) )); st=$?
[ "$st" -eq 1 ] || fail "zero multiplication must fail"
(( ((2 < 3)) )); st=$?
[ "$st" -eq 0 ] || fail "((2<3)) must succeed"
echo PASS
exit 0
