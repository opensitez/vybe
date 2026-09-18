#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_retargeting_to_another_variable
# Re-invoking 'declare -n ref=new_target' retargets the nameref to point to a new variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
slotA="A_orig"
slotB="B_orig"
declare -n pointer=slotA
pointer="A_new"
[ "$slotA" = "A_new" ] || fail "slotA modification failed"

declare -n pointer=slotB
pointer="B_new"
[ "$slotB" = "B_new" ] || fail "slotB modification failed"
[ "${!pointer}" = "slotB" ] || fail "retargeting name check failed: got [${!pointer}]"
echo PASS
exit 0
