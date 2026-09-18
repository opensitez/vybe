#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_in_subshell_does_not_affect_parent
# Executing shift inside a subshell shifts parameters in the subshell without affecting the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "p1" "p2" "p3"
(
    shift 2
    [ "$#" -eq 1 ] || exit 1
    [ "$1" = "p3" ] || exit 2
)
st=$?
[ "$st" -eq 0 ] || fail "subshell shift verification failed"
[ "$#" -eq 3 ] || fail "parent parameter count modified: want 3, got $#"
[ "$1" = "p1" ] || fail "parent param 1 altered: got [$1]"
echo PASS
exit 0
