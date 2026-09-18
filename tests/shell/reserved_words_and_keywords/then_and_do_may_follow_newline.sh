#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/then_and_do_may_follow_newline
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(
if true
then
  echo t
fi
for x in 1
do
  echo $x
done
while false
do :
done
)
[ "$out" = $'t\n1' ] || fail "got [$out]"
echo PASS
exit 0
