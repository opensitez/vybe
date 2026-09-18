#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_stderr_includes_variable_name
# The error message printed to standard error identifies the failing variable name explicitly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset UNIQUE_CONFIG_IDENTIFIER_XYZ
err_out=$( ( : "${UNIQUE_CONFIG_IDENTIFIER_XYZ:?must be defined}" ) 2>&1 )
case "$err_out" in
    *"UNIQUE_CONFIG_IDENTIFIER_XYZ"*) : ;;
    *) fail "error message missing variable identifier: got [$err_out]" ;;
esac
echo PASS
exit 0
