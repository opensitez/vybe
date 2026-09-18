#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_custom_ifs_inherited_by_subshell
# A custom IFS value configured in the parent shell is inherited and used for splitting in subshells.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IFS=':'
sub_args=$(
    (
        sample="a:b:c"
        set -- $sample
        printf '%s,%s,%s\n' "$1" "$2" "$3"
    )
)
[ "$sub_args" = "a,b,c" ] || fail "subshell failed to inherit custom IFS: got [$sub_args]"
echo PASS
exit 0
