#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/dot_and_other_regex_characters_are_literal
# The pattern is a shell glob, not a regular expression.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a.b+c'
[ "${x//./-}" = 'a-b+c' ] || fail ". is literal: got [${x//./-}]"
[ "${x//+/_}" = 'a.b_c' ] || fail "+ is literal without extglob: got [${x//+/_}]"
echo PASS
exit 0
