#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/ampersand_delimits_command_without_space
# An ampersand acts as an asynchronous command terminator even when not preceded by space.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
(exit 0)&
wait $!
st=$?
[ "$st" -eq 0 ] || fail "async wait status: want 0, got $st"
echo PASS
exit 0
