#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_pointer_pointing_to_pointer_chain
# Dereferencing chained pointers step-by-step using indirect parameter expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
actual_secret="gold"
ptr2="actual_secret"
ptr1="ptr2"
# Level 1: dereferencing ptr1 evaluates to the value of ptr2 ("actual_secret")
intermediate="${!ptr1}"
[ "$intermediate" = "actual_secret" ] || fail "first dereference failed: got [$intermediate]"
# Level 2: dereferencing intermediate evaluates to the value of actual_secret ("gold")
final="${!intermediate}"
[ "$final" = "gold" ] || fail "second dereference failed: got [$final]"
echo PASS
exit 0
