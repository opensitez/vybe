#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/negated_command_as_condition
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if ! false; then r=yes; else r=no; fi
[ "$r" = yes ] || fail "! false: got $r"
if ! [ 1 -eq 1 ]; then r=yes; else r=no; fi
[ "$r" = no ] || fail "! true test: got $r"
while ! [ "${n:-0}" -ge 2 ]; do n=$((${n:-0}+1)); done
[ "$n" -eq 2 ] || fail "negated loop condition: n=$n"
echo PASS
exit 0
