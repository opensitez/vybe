#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/ampersand_in_replacement_is_the_matched_text
# With shopt patsub_replacement (on by default) & in the replacement stands
# for the matched text; \& is a literal &; turning the option off makes & literal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s patsub_replacement
x=abc
[ "${x/b/[&]}" = 'a[b]c' ] || fail "& is the match: got [${x/b/[&]}]"
[ "${x/b/\&}" = 'a&c' ] || fail "\\& is literal: got [${x/b/\&}]"
[ "${x//[ac]/<&>}" = '<a>b<c>' ] || fail "each match: got [${x//[ac]/<&>}]"
shopt -u patsub_replacement
[ "${x/b/[&]}" = 'a[&]c' ] || fail "option off: got [${x/b/[&]}]"
echo PASS
exit 0
