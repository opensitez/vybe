#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_indexed_array_element_error
# The error substitution syntax enforces that a specific indexed array element exists.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "elem0" )
( : "${arr[5]:?element 5 is required}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "missing array element should trigger non-zero exit"

val="${arr[0]:?element 0 is required}"
[ "$val" = "elem0" ] || fail "existing element check failed: got [$val]"
echo PASS
exit 0
