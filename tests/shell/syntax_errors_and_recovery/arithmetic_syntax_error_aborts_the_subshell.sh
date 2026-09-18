#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/arithmetic_syntax_error_aborts_the_subshell
# A malformed arithmetic expression is an expansion error: in a non-interactive
# shell it aborts the current shell (here the $(…) subshell) with status 1,
# so the command after it never runs; the parent keeps going.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( { x=$((1+)); echo unreachable; } 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"operand expected"* ]] || fail "want operand expected got [$msg]"
[[ $msg != *unreachable* ]] || fail "command after the failed expansion must not run"
echo PASS
exit 0
