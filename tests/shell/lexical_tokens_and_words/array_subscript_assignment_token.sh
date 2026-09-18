#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/array_subscript_assignment_token
# An identifier with subscript [index]=value is tokenized as an assignment word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr[0]=zero
arr[2]=two
[ "${arr[0]}" = "zero" ] || fail "subscript assignment 0: want 'zero', got [${arr[0]}]"
[ "${arr[2]}" = "two" ] || fail "subscript assignment 2: want 'two', got [${arr[2]}]"
[ -z "${arr[1]}" ] || fail "unset index: want empty, got [${arr[1]}]"
echo PASS
exit 0
