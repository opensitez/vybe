#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_as_target_of_or_list
# A brace group following '||' executes only on failure and mutates surrounding variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
handled=0
false || {
    handled=1
    cause="fallback_executed"
}
[ "$handled" -eq 1 ] || fail "handled flag: want 1, got $handled"
[ "$cause" = "fallback_executed" ] || fail "cause: want 'fallback_executed', got [$cause]"
echo PASS
exit 0
