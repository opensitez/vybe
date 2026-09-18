#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_recursive_execution
# Bash functions can recursively invoke themselves and unwind call frames cleanly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
factorial() {
    local n=$1
    if [ "$n" -le 1 ]; then
        printf '1\n'
    else
        local prev
        prev=$(factorial $(( n - 1 )))
        printf '%s\n' "$(( n * prev ))"
    fi
}
fact5=$(factorial 5)
[ "$fact5" -eq 120 ] || fail "factorial 5: want 120, got $fact5"
echo PASS
exit 0
