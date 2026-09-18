#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_subshell_aborts_on_failure
# When an error substitution fails inside a subshell, the entire subshell immediately aborts.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_token
reached_end=0
(
    : "${missing_token:?abort now}"
    reached_end=1
) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "subshell should exit non-zero"
[ "$reached_end" -eq 0 ] || fail "subshell continued execution past fatal error substitution"
echo PASS
exit 0
