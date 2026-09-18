#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_export_validates_identifier
# The export builtin validates variable identifiers and exits non-zero on invalid identifier syntax.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export "1invalid_id" 2>/dev/null
st1=$?
[ "$st1" -ne 0 ] || fail "export 1invalid_id should fail"

export "bad-name" 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "export bad-name should fail"
echo PASS
exit 0
