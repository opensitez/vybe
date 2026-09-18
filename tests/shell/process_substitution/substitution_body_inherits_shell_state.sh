#!/usr/bin/env bash
# vybe-test: bash/process_substitution/substitution_body_inherits_shell_state
# The list runs in a subshell: it sees unexported variables, functions and
# positionals, but its own assignments do not come back.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- p1
v=unexported
f() { echo fn; }
read -r l < <(echo "$v $1 $(f)"; v=changed)
[ "$l" = "unexported p1 fn" ] || fail "got [$l]"
[ "$v" = unexported ] || fail "assignment leaked: v=$v"
echo PASS
exit 0
