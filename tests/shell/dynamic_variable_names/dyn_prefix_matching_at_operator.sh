#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_prefix_matching_at_operator
# The "${!prefix@}" expansion within quotes expands to individual words for each matching variable name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample_k1=1
sample_k2=2
count_matched() {
    [ "$#" -eq 2 ] || fail "matching count: want 2, got $#"
}
count_matched "${!sample_k@}"
echo PASS
exit 0
