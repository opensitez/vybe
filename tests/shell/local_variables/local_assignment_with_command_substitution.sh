#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_assignment_with_command_substitution
# Declaring 'local var=$(cmd)' captures the command substitution output directly into the local variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
capture_fn() {
    local out=$(printf 'cmd_result\n')
    [ "$out" = "cmd_result" ] || fail "local command substitution failed: got [$out]"
}
capture_fn
echo PASS
exit 0
