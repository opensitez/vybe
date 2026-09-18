#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/empty_expression_fails
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( )); st=$?
[ "$st" -eq 1 ] || fail "want 1 got $st"
e=
(( e )); st=$?
[ "$st" -eq 1 ] || fail "empty variable want 1 got $st"
echo PASS
exit 0
