#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_default_message_used_when_word_omitted
# When the error message word is omitted, Bash supplies a standard default message to stderr.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset unnamed_req_var
err_out=$( ( : "${unnamed_req_var:?}" ) 2>&1 )
case "$err_out" in
    *"parameter null or not set"*|*"unnamed_req_var"*) : ;;
    *) fail "default error message missing standard text: got [$err_out]" ;;
esac
echo PASS
exit 0
