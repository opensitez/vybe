#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_conditional_compound_body
# A function body can be a [[ ... ]] conditional expression directly without braces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
is_empty_arg() [[ -z "$1" ]]
is_empty_arg ""
st1=$?
[ "$st1" -eq 0 ] || fail "is_empty_arg '': want status 0, got $st1"

is_empty_arg "data"
st2=$?
[ "$st2" -eq 1 ] || fail "is_empty_arg 'data': want status 1, got $st2"
echo PASS
exit 0
