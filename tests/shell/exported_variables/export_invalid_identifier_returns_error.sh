#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_invalid_identifier_returns_error
# Attempting to export an invalid identifier name produces an error and non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export "1_invalid_export" 2>/dev/null
st1=$?
[ "$st1" -ne 0 ] || fail "export 1_invalid_export should return non-zero exit status"

export "hyphen-name" 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "export hyphen-name should return non-zero exit status"
echo PASS
exit 0
