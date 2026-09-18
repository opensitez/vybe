#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/nested_subshells_three_levels_deep
# Deeply nested subshells isolate environment changes at each respective depth level.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
(
    x=1
    (
        x=2
        (
            x=3
            [ "$x" -eq 3 ] || exit 1
        )
        [ "$x" -eq 2 ] || exit 1
    )
    [ "$x" -eq 1 ] || exit 1
)
[ "$x" -eq 0 ] || fail "parent x modified: got $x"
echo PASS
exit 0
