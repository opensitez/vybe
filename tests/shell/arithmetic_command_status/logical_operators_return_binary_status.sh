#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/logical_operators_return_binary_status
# Bash arithmetic logical operators normalize to 0/1 and command status is based on that.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 0 && 5 )); st=$?
[ "$st" -eq 1 ] || fail "((0 && 5)) status want 1 got $st"
(( 1 && 5 )); st=$?
[ "$st" -eq 0 ] || fail "((1 && 5)) status want 0 got $st"
(( 0 || 5 )); st=$?
[ "$st" -eq 0 ] || fail "((0 || 5)) status want 0 got $st"
(( 1 || 0 )); st=$?
[ "$st" -eq 0 ] || fail "((1 || 0)) status want 0 got $st"
(( (1 || 0) && 0 )); st=$?
[ "$st" -eq 1 ] || fail "((1 || 0) && 0) status want 1 got $st"
echo PASS
exit 0
