#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_variable_preserving_spaces_and_quotes
# Exported variables pass literal whitespace and internal quotes faithfully to child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export COMPLEX_VAL='  spaces   and "double" and '\''single'\''  '
res=$( "$BASH" -c 'printf "%s\n" "$COMPLEX_VAL"' )
expected='  spaces   and "double" and '\''single'\''  '
[ "$res" = "$expected" ] || fail "exported variable formatting corrupted: got [$res]"
echo PASS
exit 0
