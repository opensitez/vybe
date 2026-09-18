#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/command_builtin_bypasses_function_shadowing
# The 'command' builtin invokes builtins or PATH utilities, suppressing function lookup.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
printf() { builtin echo "function_intercepted"; }
fn_out=$(printf)
cmd_out=$(command printf '%s' "direct_printf")
unset -f printf
[ "$fn_out" = "function_intercepted" ] || fail "function out: got [$fn_out]"
[ "$cmd_out" = "direct_printf" ] || fail "command printf: got [$cmd_out]"
echo PASS
exit 0
