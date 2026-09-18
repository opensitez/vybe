#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/no_parameter_expansion_inside
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
v=value
x=$'$v ${v} $(echo n) `echo n`'
[ "$x" = '$v ${v} $(echo n) `echo n`' ] || fail "got [$x]"
echo PASS
exit 0
