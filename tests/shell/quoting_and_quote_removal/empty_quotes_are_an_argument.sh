#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/empty_quotes_are_an_argument
# "" and '' each produce a (null) argument, unlike an unquoted expansion of
# an empty variable, which produces none.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
empty=
n=$(count "" '')
[ "$n" = 2 ] || fail "quoted empties: want 2 got $n"
n=$(count $empty)
[ "$n" = 0 ] || fail "unquoted empty expansion: want 0 got $n"
n=$(count "$empty")
[ "$n" = 1 ] || fail "quoted empty expansion: want 1 got $n"
echo PASS
exit 0
