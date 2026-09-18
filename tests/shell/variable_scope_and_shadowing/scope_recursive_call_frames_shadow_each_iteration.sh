#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_recursive_call_frames_shadow_each_iteration
# Recursive function invocations instantiate distinct local scopes that restore sequentially as frames unwind.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
audit=""
recurse_scope() {
    local depth=$1
    local frame="F$depth"
    if [ "$depth" -lt 3 ]; then
        recurse_scope $(( depth + 1 ))
    fi
    audit+="$frame,"
}
recurse_scope 1
[ "$audit" = "F3,F2,F1," ] || fail "recursive unwinding corrupted scope frames: got [$audit]"
echo PASS
exit 0
