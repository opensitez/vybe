#!/usr/bin/env bash
# vybe-test: bash/double_bracket_conditionals/bare_string_is_a_nonempty_test
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
e=; x=0
[[ 0 ]] || fail "the string 0 is non-empty"
[[ $x ]] || fail "variable holding 0"
[[ "" ]] && fail "empty literal"
[[ $e ]] && fail "empty variable, unquoted, is fine inside [["
[[ ! $e ]] || fail "negated empty"
echo PASS
exit 0
