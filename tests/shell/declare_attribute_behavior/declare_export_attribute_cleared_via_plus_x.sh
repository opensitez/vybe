#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_export_attribute_cleared_via_plus_x
# The 'declare +x' flag unexports a variable, preventing it from being passed to child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -x PRIVATE_DATA="exported_first"
declare +x PRIVATE_DATA
res=$(
    "$BASH" -c 'printf "%s\n" "$PRIVATE_DATA"'
)
[ -z "$res" ] || fail "declare +x failed to unexport variable: child saw [$res]"
[ "$PRIVATE_DATA" = "exported_first" ] || fail "local variable value lost after +x: got [$PRIVATE_DATA]"
echo PASS
exit 0
