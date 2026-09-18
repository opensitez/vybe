#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/plus_equals_assignment_token
# The += operator is recognized as an append-assignment token at the start of a command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val="hello"
val+=" world"
[ "$val" = "hello world" ] || fail "plus-equals token: want 'hello world', got [$val]"
echo PASS
exit 0
