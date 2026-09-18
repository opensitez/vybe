#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_subshell_mutation_isolated_from_parent
# Changes made to positional parameters inside a subshell do not affect the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "parent_a" "parent_b"
(
    set -- "sub_only_1" "sub_only_2" "sub_only_3"
    [ "$#" -eq 3 ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell set failed"
[ "$#" -eq 2 ] || fail "parent parameter count modified: want 2, got $#"
[ "$1" = "parent_a" ] && [ "$2" = "parent_b" ] || fail "parent parameters altered"
echo PASS
exit 0
