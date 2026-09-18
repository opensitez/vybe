#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/builtin_command_bypasses_function_shadowing
# The 'builtin' builtin executes the shell builtin directly, bypassing any function shadow.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
echo() { printf 'custom_echo\n'; }
fn_res=$(echo)
builtin_res=$(builtin echo "real_builtin")
unset -f echo
[ "$fn_res" = "custom_echo" ] || fail "function invocation: got [$fn_res]"
[ "$builtin_res" = "real_builtin" ] || fail "builtin echo bypass: got [$builtin_res]"
echo PASS
exit 0
