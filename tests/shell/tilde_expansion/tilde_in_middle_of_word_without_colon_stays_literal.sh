#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_middle_of_word_without_colon_stays_literal
# A tilde occurring within the body of a word (not at the start, and not after a colon in assignment) is literal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
HOME="/custom/home"
res=prefix~suffix
[ "$res" = "prefix~suffix" ] || fail "mid-word tilde was improperly expanded: got [$res]"
res2=foo~/bar
[ "$res2" = "foo~/bar" ] || fail "mid-word foo~/bar was improperly expanded: got [$res2]"
echo PASS
exit 0
