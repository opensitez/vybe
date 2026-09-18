#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/special_parameter_exit_status_resolution
# The special parameter identifier '?' resolves to the numeric exit status of the last executed command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(exit 0)
st0=$?
[ "$st0" -eq 0 ] || fail "st0: want 0, got $st0"

(exit 42)
st42=$?
[ "$st42" -eq 42 ] || fail "st42: want 42, got $st42"
echo PASS
exit 0
