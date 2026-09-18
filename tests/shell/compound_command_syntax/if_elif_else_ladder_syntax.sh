#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/if_elif_else_ladder_syntax
# The if ... elif ... else ... fi compound command executes precisely one matching branch.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
classify() {
    local val="$1"
    if [ "$val" -lt 0 ]; then
        printf 'negative\n'
    elif [ "$val" -eq 0 ]; then
        printf 'zero\n'
    else
        printf 'positive\n'
    fi
}
[ "$(classify -5)" = "negative" ] || fail "classify -5: want 'negative'"
[ "$(classify 0)" = "zero" ] || fail "classify 0: want 'zero'"
[ "$(classify 10)" = "positive" ] || fail "classify 10: want 'positive'"
echo PASS
exit 0
