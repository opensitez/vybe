#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/readonly_assignment_is_rejected_during_lookup_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly a=9
(( a = a + 1 )) 2>/dev/null; st=$?
[ "$st" -eq 1 ] || fail "readonly assignment should fail"
[ "$a" -eq 9 ] || fail "readonly variable unchanged"
echo PASS
exit 0
