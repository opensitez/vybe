#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_shadowing_integer_global_with_string_local
# A local variable declared without the integer attribute shadows an integer global variable as text.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i num=50
test_type_shadow() {
    local num="hello world"
    [ "$num" = "hello world" ] || fail "local text shadow failed: got [$num]"
}
test_type_shadow
[ "$num" -eq 50 ] || fail "global integer corrupted: got $num"
echo PASS
exit 0
