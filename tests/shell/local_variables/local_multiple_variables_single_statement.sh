#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_multiple_variables_single_statement
# Multiple local variables can be declared and initialized in a single 'local' statement.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multi_local_fn() {
    local a=10 b=20 c=30
    [ "$a" -eq 10 ] && [ "$b" -eq 20 ] && [ "$c" -eq 30 ] || exit 1
}
multi_local_fn
st=$?
[ "$st" -eq 0 ] || fail "multiple local variable declaration failed"
echo PASS
exit 0
