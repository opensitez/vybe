#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/unquoted_right_side_is_a_pattern_quoted_is_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
[[ $x == a* ]] || fail "unquoted a* must match"
[[ $x == "a*" ]] && fail "quoted a* must be literal"
[[ 'a*' == "a*" ]] || fail "literal compare of a* with itself"
echo PASS
exit 0
