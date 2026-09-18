#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_builtin_name_identifier_allowed
# Builtin command names (such as 'echo', 'read', 'cd') can be variable names without masking commands.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
echo="message_payload"
read="input_data"
out=$(echo "$echo")
[ "$out" = "message_payload" ] || fail "echo variable masked echo builtin: got [$out]"
[ "$read" = "input_data" ] || fail "read variable assignment failed"
echo PASS
exit 0
