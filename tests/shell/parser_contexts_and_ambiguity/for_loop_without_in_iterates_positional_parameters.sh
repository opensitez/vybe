#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/for_loop_without_in_iterates_positional_parameters
# Omitting 'in words' in a for loop syntax causes iteration over positional parameters ($@).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
test_iteration() {
    local acc=""
    for item; do
        acc+="${item}_"
    done
    [ "$acc" = "p1_p2_p3_" ] || fail "positional iteration: want 'p1_p2_p3_', got [$acc]"
}
test_iteration p1 p2 p3
echo PASS
exit 0
