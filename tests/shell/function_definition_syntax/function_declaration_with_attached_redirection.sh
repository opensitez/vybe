#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_declaration_with_attached_redirection
# Redirections specified on a function definition apply to all executions of that function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
redirected_fn() {
    printf 'routed_message\n'
} >&2
captured_err=$(
    redirected_fn 2>&1 >/dev/null
)
[ "$captured_err" = "routed_message" ] || fail "attached redirection failed: got [$captured_err]"
echo PASS
exit 0
