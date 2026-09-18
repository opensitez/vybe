#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_stderr_redirection_aggregates_all_stderr
# Redirecting 2> on a brace group captures standard error from all enclosed commands.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
err_out=$(
    {
        printf 'err_msg_1\n' >&2
        printf 'err_msg_2\n' >&2
    } 2>&1 >/dev/null
)
expected=$(printf 'err_msg_1\nerr_msg_2')
[ "$err_out" = "$expected" ] || fail "aggregated stderr: got [$err_out]"
echo PASS
exit 0
