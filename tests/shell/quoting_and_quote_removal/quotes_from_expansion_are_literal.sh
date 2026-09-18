#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quotes_from_expansion_are_literal
# Quote removal applies to quotes in the source text, not to quote characters
# that come out of an expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
second() { echo "$2"; }
x="'a b'"
n=$(count $x)
[ "$n" = 2 ] || fail "want 2 args (quotes do not group), got $n"
v=$(second $x)
[ "$v" = "b'" ] || fail "want [b'] got [$v]"
echo PASS
exit 0
