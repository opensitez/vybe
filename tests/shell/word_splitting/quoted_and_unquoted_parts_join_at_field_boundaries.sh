#!/usr/bin/env bash
# vybe-test: bash/word_splitting/quoted_and_unquoted_parts_join_at_field_boundaries
# Only the unquoted expansion is split; its first field glues to the quoted
# text before it and its last field to the text after it.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='c d'
args() { printf '[%s]' "$@"; echo; }
out=$(args "a b"$x"e f")
[ "$out" = '[a bc][de f]' ] || fail "got [$out]"
echo PASS
exit 0
