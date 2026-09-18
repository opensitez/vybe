#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_declare_dash_g_mutates_global_ignoring_local_shadow
# The 'declare -g var=val' statement mutates the global variable even when a local shadow exists.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="initial_global"
mutate_global_fn() {
    local x="local_shadow"
    declare -g x="mutated_global"
    [ "$x" = "local_shadow" ] || fail "local shadow was altered by declare -g: got [$x]"
}
mutate_global_fn
[ "$x" = "mutated_global" ] || fail "declare -g failed to mutate global variable: got [$x]"
echo PASS
exit 0
