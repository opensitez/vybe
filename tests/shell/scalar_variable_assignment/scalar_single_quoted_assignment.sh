#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_single_quoted_assignment
# Single quotes on RHS preserve all enclosed characters literally without interpolation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="test"
var='$x and `date` and "quoted"'
[ "$var" = '$x and `date` and "quoted"' ] || fail "single quote assignment failed: got [$var]"
echo PASS
exit 0
