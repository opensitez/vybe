#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/type_reports_reserved_words_as_keyword
# [[ is a keyword while [ is a builtin.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
for w in if then else elif fi for while until do done case esac function select time '[[' ']]' '{' '}' '!' coproc; do
  t=$(type -t "$w")
  [ "$t" = keyword ] || fail "$w: want keyword got [$t]"
done
t=$(type -t '[')
[ "$t" = builtin ] || fail "[ want builtin got [$t]"
echo PASS
exit 0
