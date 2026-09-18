#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/nested_arithmetic_expansions
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $(( $((1+1)) * 3 )) -eq 6 ] || fail "got $(( $((1+1)) * 3 ))"
n=$(( ${#HOME} > 0 ? 1 : 0 ))
[ "$n" -eq 1 ] || fail "parameter expansion inside arithmetic: got $n"
echo PASS
exit 0
