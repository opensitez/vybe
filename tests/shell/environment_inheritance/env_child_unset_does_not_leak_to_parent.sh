#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_child_unset_does_not_leak_to_parent
# Unsetting an inherited environment variable in a child process does not unset it in the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export RESILIENT_KEY="present_in_parent"
"$BASH" -c 'unset RESILIENT_KEY; [[ ! -v RESILIENT_KEY ]] || exit 1'
st=$?
[ "$st" -eq 0 ] || fail "child unset failed"
[ "$RESILIENT_KEY" = "present_in_parent" ] || fail "parent variable was removed by child unset"
echo PASS
exit 0
