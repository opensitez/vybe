#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/adjacent_quoted_segments_form_one_argument
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
first() { echo "$1"; }
count() { echo $#; }
n=$(count 'a'"b"c)
[ "$n" = 1 ] || fail "want 1 arg got $n"
v=$(first 'a'"b"c)
[ "$v" = abc ] || fail "want [abc] got [$v]"
echo PASS
exit 0
