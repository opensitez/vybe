#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/quoted_part_of_regex_is_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc123
[[ $x =~ b.1 ]] || fail "unquoted . is a metacharacter"
[[ $x =~ "b.1" ]] && fail "quoted b.1 is literal and must not match"
[[ "a.b" =~ a\.b ]] || fail "escaped dot matches a dot"
[[ axb =~ a\.b ]] && fail "escaped dot must not match x"
echo PASS
exit 0
