#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/assignment_after_command_name_is_word
# name=value is an assignment only before the command name; after it, it is
# an ordinary argument word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset a
out=$(echo a=b)
[ "$out" = "a=b" ] || fail "want [a=b] got [$out]"
[ -z "${a+set}" ] || fail "a must not have been assigned"
echo PASS
exit 0
