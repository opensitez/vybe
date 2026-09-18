#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/export_command_preserves_unquoted_spaces_in_assignment
# Arguments to declaration commands like export and readonly are parsed as assignments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export EXP_A="hello world" EXP_B=singleword
[ "$EXP_A" = "hello world" ] || fail "EXP_A: want 'hello world', got [$EXP_A]"
[ "$EXP_B" = "singleword" ] || fail "EXP_B: want 'singleword', got [$EXP_B]"
echo PASS
exit 0
