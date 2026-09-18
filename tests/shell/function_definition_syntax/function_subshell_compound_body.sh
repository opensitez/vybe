#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_subshell_compound_body
# A function body can be a subshell ( ... ) instead of a brace group, isolating internal changes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10
subshell_body_fn() (
    x=99
    cd /tmp
)
subshell_body_fn
[ "$x" -eq 10 ] || fail "function subshell mutated variable x: got $x"
[ "$PWD" != "/tmp" ] || fail "function subshell mutated PWD"
echo PASS
exit 0
