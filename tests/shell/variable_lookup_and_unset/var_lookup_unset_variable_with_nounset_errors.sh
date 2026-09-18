#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_unset_variable_with_nounset_errors
# Referencing an unset variable when 'set -u' (nounset) is enabled aborts execution with an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
    set -u
    unset undeclared_lookup
    echo "$undeclared_lookup" 2>/dev/null
    exit 0
)
st=$?
[ "$st" -ne 0 ] || fail "nounset should abort on referencing unset variable"
echo PASS
exit 0
