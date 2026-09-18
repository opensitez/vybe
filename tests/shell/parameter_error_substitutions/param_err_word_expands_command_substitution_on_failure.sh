#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_word_expands_command_substitution_on_failure
# When the error condition is triggered, command substitutions inside the error word are executed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset trigger_var
err_out=$( ( : "${trigger_var:?$(printf 'dynamic_failure_message\n')}" ) 2>&1 )
case "$err_out" in
    *"dynamic_failure_message"*) : ;;
    *) fail "command substitution in error message failed: got [$err_out]" ;;
esac
echo PASS
exit 0
