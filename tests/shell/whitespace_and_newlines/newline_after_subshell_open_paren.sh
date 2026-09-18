#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_after_subshell_open_paren
# A newline after an open parenthesis '(' within a compound command continues the subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=$(
    echo "inside"
)
[ "$res" = "inside" ] || fail "newline after subshell open paren: want 'inside', got [$res]"
echo PASS
exit 0
