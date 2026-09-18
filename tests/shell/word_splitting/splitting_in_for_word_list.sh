#!/usr/bin/env bash
# vybe-test: bash/word_splitting/splitting_in_for_word_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='one two three'
n=0; for w in $x; do n=$((n+1)); last=$w; done
[ "$n" -eq 3 ] && [ "$last" = three ] || fail "unquoted: n=$n last=$last"
n=0; for w in "$x"; do n=$((n+1)); done
[ "$n" -eq 1 ] || fail "quoted: want 1 got $n"
echo PASS
exit 0
