#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_after_pipe_continues_command
# A newline immediately following a pipe operator does not terminate the command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=$(printf '%s' "hello" |
    cat)
[ "$res" = "hello" ] || fail "newline after pipe: want 'hello', got [$res]"
echo PASS
exit 0
