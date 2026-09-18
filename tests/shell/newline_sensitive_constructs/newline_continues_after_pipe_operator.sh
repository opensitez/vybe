#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_continues_after_pipe_operator
# A newline immediately following a pipe operator continues the pipeline to the next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(printf 'pipe_token\n' |
    cat)
[ "$out" = "pipe_token" ] || fail "newline after pipe: got [$out]"
echo PASS
exit 0
