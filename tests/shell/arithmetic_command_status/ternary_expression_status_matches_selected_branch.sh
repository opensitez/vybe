#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/ternary_expression_status_matches_selected_branch
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 1 ? 7 : 0 )); st=$?
[ "$st" -eq 0 ] || fail "true branch returns 7"
(( 0 ? 0 : 9 )); st=$?
[ "$st" -eq 0 ] || fail "false branch returns 9"
(( 0 ? 0 : 0 )); st=$?
[ "$st" -eq 1 ] || fail "branch returning 0 should fail"
echo PASS
exit 0
