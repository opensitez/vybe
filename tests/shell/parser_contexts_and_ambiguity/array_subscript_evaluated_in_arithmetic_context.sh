#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/array_subscript_evaluated_in_arithmetic_context
# The index inside array subscript brackets arr[...] is evaluated in an arithmetic context.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=(zero one two three four five)
idx=2
val="${arr[idx + 1]}"
[ "$val" = "three" ] || fail "arithmetic in subscript: want 'three', got [$val]"
echo PASS
exit 0
