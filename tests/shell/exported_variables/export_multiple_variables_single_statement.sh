#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_multiple_variables_single_statement
# Multiple variables can be initialized and marked for export in a single export statement.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export E1="val1" E2="val2" E3="val3"
res=$( "$BASH" -c 'printf "%s,%s,%s\n" "$E1" "$E2" "$E3"' )
[ "$res" = "val1,val2,val3" ] || fail "multiple export failed: got [$res]"
echo PASS
exit 0
