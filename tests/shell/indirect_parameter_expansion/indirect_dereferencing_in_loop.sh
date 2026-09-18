#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_dereferencing_in_loop
# Indirect parameter expansion operates reliably when stepping through variable names in a loop.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var_a="val1"
var_b="val2"
var_c="val3"
collected=""
for name in var_a var_b var_c; do
    collected+="${!name},"
done
[ "$collected" = "val1,val2,val3," ] || fail "loop dereferencing failed: got [$collected]"
echo PASS
exit 0
