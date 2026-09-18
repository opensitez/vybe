#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_global_flag_dash_g_inside_function
# The 'declare -g' flag forces variable creation in the global scope when invoked inside a function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset GLOBAL_CREATED
creator_fn() {
    declare -g GLOBAL_CREATED="global_payload"
}
creator_fn
[ "$GLOBAL_CREATED" = "global_payload" ] || fail "declare -g inside function failed to create global variable"
echo PASS
exit 0
