#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_return_from_nested_control_structure
# Executing 'return' inside nested loops or conditionals terminates the entire function immediately.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
search_item() {
    local target=$1
    for num in 10 20 30 40; do
        if [ "$num" -eq "$target" ]; then
            return 0
        fi
    done
    return 1
}
search_item 20
[ "$?" -eq 0 ] || fail "search_item 20 should find item and return 0"

search_item 99
[ "$?" -eq 1 ] || fail "search_item 99 should fail and return 1"
echo PASS
exit 0
