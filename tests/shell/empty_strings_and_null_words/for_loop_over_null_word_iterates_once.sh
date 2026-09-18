#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/for_loop_over_null_word_iterates_once
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
e=
n=0; for w in ""; do n=$((n+1)); done
[ "$n" -eq 1 ] || fail "quoted empty: want 1 iteration got $n"
n=0; for w in $e; do n=$((n+1)); done
[ "$n" -eq 0 ] || fail "unquoted empty expansion: want 0 iterations got $n"
n=0; for w in "$e" "$e"; do n=$((n+1)); done
[ "$n" -eq 2 ] || fail "two quoted empties: want 2 got $n"
echo PASS
exit 0
