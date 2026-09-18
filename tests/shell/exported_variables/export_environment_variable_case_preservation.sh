#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_environment_variable_case_preservation
# Environment variable names maintain strict case-sensitivity when exported to child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export TargetCase="camel"
export TARGETCASE="upper"
export targetcase="lower"
res=$( "$BASH" -c 'printf "%s,%s,%s\n" "$TargetCase" "$TARGETCASE" "$targetcase"' )
[ "$res" = "camel,upper,lower" ] || fail "case-sensitivity corrupted across export: got [$res]"
echo PASS
exit 0
