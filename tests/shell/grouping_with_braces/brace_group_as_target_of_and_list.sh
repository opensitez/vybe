#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_as_target_of_and_list
# A brace group following '&&' executes on success and allows multi-statement sequential setup.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ran=0
true && {
    ran=1
    tag="and_success"
}
[ "$ran" -eq 1 ] || fail "ran: want 1, got $ran"
[ "$tag" = "and_success" ] || fail "tag: want 'and_success', got [$tag]"
echo PASS
exit 0
