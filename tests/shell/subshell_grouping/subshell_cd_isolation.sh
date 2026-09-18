#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_cd_isolation
# Running 'cd' inside a subshell changes directory only for that subshell, leaving parent $PWD intact.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
parent_dir="$PWD"
(
    cd /tmp
    [ "$PWD" = "/tmp" ] || [ "$PWD" = "/private/tmp" ] || exit 1
)
[ "$PWD" = "$parent_dir" ] || fail "parent PWD changed: want [$parent_dir], got [$PWD]"
echo PASS
exit 0
