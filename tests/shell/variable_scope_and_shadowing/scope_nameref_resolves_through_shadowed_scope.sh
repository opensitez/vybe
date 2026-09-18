#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_nameref_resolves_through_shadowed_scope
# A nameref resolves dynamically to the nearest shadowed variable in the call stack.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read_ref() {
    local -n ref=$1
    printf '%s\n' "$ref"
}
caller_a() {
    local target="caller_a_value"
    read_ref target
}
caller_b() {
    local target="caller_b_value"
    read_ref target
}
target="global_target_value"
out_a=$(caller_a)
out_b=$(caller_b)
out_glob=$(read_ref target)
[ "$out_a" = "caller_a_value" ] || fail "caller_a nameref failed: got [$out_a]"
[ "$out_b" = "caller_b_value" ] || fail "caller_b nameref failed: got [$out_b]"
[ "$out_glob" = "global_target_value" ] || fail "global nameref failed: got [$out_glob]"
echo PASS
exit 0
