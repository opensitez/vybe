#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/negation_applies_to_one_pipeline_not_the_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
! true && false; st=$?
[ "$st" -eq 1 ] || fail "(! true) && false: want 1 got $st"
! { true && false; }; st=$?
[ "$st" -eq 0 ] || fail "! { true && false; }: want 0 got $st"
echo PASS
exit 0
