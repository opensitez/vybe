#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_implicit_for_loop_iteration
# A for loop without an 'in' clause iterates implicitly over positional parameters "$@".
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "p1" "p2" "p3"
accum=""
for arg; do
    accum+="$arg;"
done
[ "$accum" = "p1;p2;p3;" ] || fail "implicit for loop over positional parameters failed: got [$accum]"
echo PASS
exit 0
