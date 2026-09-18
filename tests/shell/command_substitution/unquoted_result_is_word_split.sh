#!/usr/bin/env bash
# vybe-test: bash/command_substitution/unquoted_result_is_word_split
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
[ "$(count $(echo "a b  c"))" = 3 ] || fail "unquoted: want 3"
[ "$(count "$(echo "a b  c")")" = 1 ] || fail "quoted: want 1"
[ "$(count $(true))" = 0 ] || fail "empty unquoted vanishes"
[ "$(count $(printf 'x\ny'))" = 2 ] || fail "newlines split too"
echo PASS
exit 0
