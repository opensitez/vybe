#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/comparison_with_assignment_status_in_if
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
if (( x = 0 )); then then='bad'; else then='good'; fi
[ "$then" = good ] || fail "assignment to 0 should fail in if"
x=4
if (( x = 3, x < 10 )); then then='good'; else then='bad'; fi
[ "$then" = good ] || fail "comma and relational branch"
[ "$x" -eq 3 ] || fail "x must remain 3"
echo PASS
exit 0
