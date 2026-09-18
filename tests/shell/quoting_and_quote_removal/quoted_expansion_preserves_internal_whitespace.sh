#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quoted_expansion_preserves_internal_whitespace
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
x="a  b   c"
n=$(count "$x")
[ "$n" = 1 ] || fail "quoted: want 1 got $n"
n=$(count $x)
[ "$n" = 3 ] || fail "unquoted: want 3 got $n"
first() { echo "$1"; }
[ "$(first "$x")" = "a  b   c" ] || fail "runs of blanks must survive quoting"
echo PASS
exit 0
