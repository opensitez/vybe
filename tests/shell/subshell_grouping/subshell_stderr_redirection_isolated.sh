#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_stderr_redirection_isolated
# Standard error from all commands in a subshell can be redirected to a dedicated stream.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
err_data=$(
    (
        printf 'sub_err_1\n' >&2
        printf 'sub_err_2\n' >&2
    ) 2>&1 >/dev/null
)
expected=$(printf 'sub_err_1\nsub_err_2')
[ "$err_data" = "$expected" ] || fail "subshell stderr redirection: got [$err_data]"
echo PASS
exit 0
