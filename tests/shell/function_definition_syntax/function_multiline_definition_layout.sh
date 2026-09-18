#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_multiline_definition_layout
# Functions can be formatted across newlines with the opening brace on a separate line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
block_fn()
{
    step1="a"
    step2="b"
    printf '%s%s\n' "$step1" "$step2"
}
res=$(block_fn)
[ "$res" = "ab" ] || fail "multiline function layout failed: got [$res]"
echo PASS
exit 0
