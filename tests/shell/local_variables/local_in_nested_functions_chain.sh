#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_in_nested_functions_chain
# Nested function definitions within functions maintain proper dynamic scoping across the chain.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
outer_scope() {
    local depth="outer"
    inner_scope() {
        [ "$depth" = "outer" ] || fail "nested function lookup failed: got [$depth]"
        local depth="inner"
        [ "$depth" = "inner" ] || fail "nested local failed: got [$depth]"
    }
    inner_scope
    [ "$depth" = "outer" ] || fail "outer local corrupted by nested function: got [$depth]"
}
outer_scope
echo PASS
exit 0
