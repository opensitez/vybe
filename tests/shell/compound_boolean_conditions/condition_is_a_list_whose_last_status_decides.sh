#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/condition_is_a_list_whose_last_status_decides
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if false; true; then r=a; else r=b; fi
[ "$r" = a ] || fail "false; true -> true branch, got $r"
if true; false; then r=a; else r=b; fi
[ "$r" = b ] || fail "true; false -> else branch, got $r"
if { false; true; }; then r=a; else r=b; fi
[ "$r" = a ] || fail "brace group condition: got $r"
echo PASS
exit 0
