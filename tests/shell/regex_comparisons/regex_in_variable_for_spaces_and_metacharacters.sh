#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/regex_in_variable_for_spaces_and_metacharacters
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
re='^a b$'
[[ "a b" =~ $re ]] || fail "space via variable"
re='^[0-9]{3}-[0-9]{4}$'
[[ 555-1234 =~ $re ]] || fail "phone pattern"
[[ 55-1234 =~ $re ]] && fail "too few digits must not match"
[[ "a b" =~ "$re" ]] && fail "quoted variable is literal text"
echo PASS
exit 0
