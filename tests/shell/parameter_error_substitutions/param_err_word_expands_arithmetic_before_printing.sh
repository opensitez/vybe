#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_word_expands_arithmetic_before_printing
# The error message undergoes arithmetic expansion before being emitted to standard error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_metric
err_out=$( ( : "${missing_metric:?failed with code $(( 400 + 4 ))}" ) 2>&1 )
case "$err_out" in
    *"failed with code 404"*) : ;;
    *) fail "arithmetic in error word failed: got [$err_out]" ;;
esac
echo PASS
exit 0
