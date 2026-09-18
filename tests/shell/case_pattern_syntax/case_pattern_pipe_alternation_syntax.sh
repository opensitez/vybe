#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_pipe_alternation_syntax
# Multiple alternative patterns can be combined with '|' inside a single case clause.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval_code() {
    case "$1" in
        200|201|204) printf 'success\n' ;;
        400|404) printf 'client_error\n' ;;
        500|503) printf 'server_error\n' ;;
        *) printf 'unknown\n' ;;
    esac
}
[ "$(eval_code 201)" = "success" ] || fail "201 match failed"
[ "$(eval_code 404)" = "client_error" ] || fail "404 match failed"
[ "$(eval_code 503)" = "server_error" ] || fail "503 match failed"
[ "$(eval_code 999)" = "unknown" ] || fail "999 match failed"
echo PASS
exit 0
