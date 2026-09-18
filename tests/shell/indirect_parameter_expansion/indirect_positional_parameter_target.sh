#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_positional_parameter_target
# The ${!ptr} expansion where ptr="1" retrieves positional parameter $1.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "pos1_val" "pos2_val"
ptr="1"
[ "${!ptr}" = "pos1_val" ] || fail "indirect \$1 failed: got [${!ptr}]"
ptr="2"
[ "${!ptr}" = "pos2_val" ] || fail "indirect \$2 failed: got [${!ptr}]"
echo PASS
exit 0
