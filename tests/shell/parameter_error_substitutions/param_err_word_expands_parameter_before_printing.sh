#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_word_expands_parameter_before_printing
# The error message undergoes parameter expansion before being written to standard error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
SERVICE="auth_server"
unset PORT
err_out=$( ( : "${PORT:?service $SERVICE requires a port}" ) 2>&1 )
case "$err_out" in
    *"service auth_server requires a port"*) : ;;
    *) fail "parameter expansion in error message failed: got [$err_out]" ;;
esac
echo PASS
exit 0
