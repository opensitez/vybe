#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/string_operators_are_lexical_not_numeric
# "10" sorts before "9" as a string; only -lt and (( )) compare numbers.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 10 < 9 ]] || fail "[[ 10 < 9 ]] must be true lexically"
[ 10 \< 9 ] || fail "[ 10 \\< 9 ] must be true lexically"
[[ 10 -lt 9 ]] && fail "[[ 10 -lt 9 ]] must be false"
(( 10 < 9 )) && fail "(( 10 < 9 )) must be false"
echo PASS
exit 0
