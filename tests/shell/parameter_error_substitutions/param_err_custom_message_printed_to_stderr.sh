#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_custom_message_printed_to_stderr
# The custom error message supplied to :? is emitted to standard error upon failure.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_database_url
err_out=$( ( : "${missing_database_url:?DATABASE_URL is required to boot}" ) 2>&1 )
case "$err_out" in
    *"DATABASE_URL is required to boot"*) : ;;
    *) fail "stderr missing custom error message: got [$err_out]" ;;
esac
echo PASS
exit 0
