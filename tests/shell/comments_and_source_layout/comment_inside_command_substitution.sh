#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_inside_command_substitution
# Comments inside a $(...) command substitution are ignored and do not affect the captured output.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val=$(
    # calculate result
    echo 42 # return value
)
[ "$val" = "42" ] || fail "comment inside cmdsub: want '42', got [$val]"
echo PASS
exit 0
