#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/at_u_capitalizes_only_first_letter
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="hello World"
[ "${x@u}" = "Hello World" ] || fail "got [${x@u}]"
y="1abc"
[ "${y@u}" = "1abc" ] || fail "leading digit: got [${y@u}]"
echo PASS
exit 0
