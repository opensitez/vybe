#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_prefix_matching_at_operator_quoted
# The "${!prefix@}" expansion inside double quotes expands to distinct arguments for each matching variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
audit_k1="v1"
audit_k2="v2"
audit_k3="v3"
count_matches() {
    [ "$#" -eq 3 ] || fail "matching count: want 3, got $#"
}
count_matches "${!audit_k@}"
echo PASS
exit 0
