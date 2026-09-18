#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_special_param_hash_target
# The ${!ptr} expansion where ptr="#" retrieves the count of positional parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "x" "y" "z"
ptr="#"
val="${!ptr}"
[ "$val" -eq 3 ] || fail "indirect \$# resolution failed: want 3, got [$val]"
echo PASS
exit 0
