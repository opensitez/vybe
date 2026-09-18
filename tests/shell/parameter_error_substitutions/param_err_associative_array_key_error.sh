#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_associative_array_key_error
# The error substitution syntax enforces that a specific associative array key is present.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A config=( [host]="127.0.0.1" )
( : "${config[port]:?port configuration is mandatory}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "missing associative array key should trigger non-zero exit"

val="${config[host]:?host is required}"
[ "$val" = "127.0.0.1" ] || fail "present key check failed: got [$val]"
echo PASS
exit 0
