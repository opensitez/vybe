#!/usr/bin/env bash
# vybe-test: bash/word_splitting/no_splitting_of_case_word
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a b'
case $x in
  "a b") ok=yes ;;
  *) ok=no ;;
esac
[ "$ok" = yes ] || fail "case word was split"
echo PASS
exit 0
