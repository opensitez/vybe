#!/usr/bin/env bash
# vybe-test: bash/string_comparison_operators/lexical_ordering_in_double_bracket_and_escaped_in_single
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ a < b ]] && [[ b > a ]] || fail "[[ ordering"
[ a \< b ] && [ b \> a ] || fail "[ ordering needs escaped operators"
[[ a < a ]] && fail "equal strings are not less"
[[ a > a ]] && fail "equal strings are not greater"
[[ "a " == "a" ]] && fail "trailing space is significant"
echo PASS
exit 0
