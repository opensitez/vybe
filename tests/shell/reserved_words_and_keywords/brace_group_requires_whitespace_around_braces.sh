#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/brace_group_requires_whitespace_around_braces
# { and } are reserved words, not metacharacters, so {cmd;} without space fails parsing.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '{echo 1;}' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "braces without whitespace should fail as unknown command"

{ echo 1; } >/dev/null
st=$?
[ "$st" -eq 0 ] || fail "braces with whitespace: want 0, got $st"
echo PASS
exit 0
