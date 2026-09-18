#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_dash_F_lists_function_names
# The 'declare -F func_name' outputs just the function name and returns status 0 if defined.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample_target_fn() { :; }
out=$(declare -F sample_target_fn)
[ "$out" = "sample_target_fn" ] || fail "declare -F output: want 'sample_target_fn', got [$out]"

declare -F unfindable_func_xyz 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "declare -F on undefined function should return non-zero exit code"
echo PASS
exit 0
