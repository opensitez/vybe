#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_variable_value_accessible_in_lookups
# A readonly variable behaves as a standard variable in all parameter expansion contexts.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly PATH_PREFIX="/usr/local"
[ "$PATH_PREFIX/bin" = "/usr/local/bin" ] || fail "interpolation failed"
[ "${#PATH_PREFIX}" -eq 10 ] || fail "length expansion failed: want 10, got ${#PATH_PREFIX}"
[ "${PATH_PREFIX:5}" = "local" ] || fail "substring expansion failed: want 'local', got [${PATH_PREFIX:5}]"
echo PASS
exit 0
