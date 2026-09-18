#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/single_quote_via_close_escape_reopen
# A single quote cannot appear inside single quotes; 'it'\''s' is the idiom.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
v='it'\''s'
[ "$v" = "it's" ] || fail "want [it's] got [$v]"
n=$(count 'it'\''s')
[ "$n" = 1 ] || fail "must stay one word, got $n"
echo PASS
exit 0
