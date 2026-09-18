#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quoted_glob_characters_stay_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
n=$(count "*" '?' "[a]")
[ "$n" = 3 ] || fail "want 3 got $n"
out=$(echo "*" '?' "[a]")
[ "$out" = '* ? [a]' ] || fail "want [* ? [a]] got [$out]"
echo PASS
exit 0
