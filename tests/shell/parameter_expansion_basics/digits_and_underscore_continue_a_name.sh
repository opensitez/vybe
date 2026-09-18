#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/digits_and_underscore_continue_a_name
# $x1 reads the variable x1, not x followed by 1; braces are needed to stop
# the name early.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=v; x1=w; x_y=z
out=$(echo $x1 ${x}1 $x_y ${x}_y)
[ "$out" = 'w v1 z v_y' ] || fail "got [$out]"
echo PASS
exit 0
