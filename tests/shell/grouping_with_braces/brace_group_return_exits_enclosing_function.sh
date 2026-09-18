#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_return_exits_enclosing_function
# A 'return' executed inside a brace group within a function immediately returns from that function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fn_test() {
    {
        return 42
    }
    printf 'unreachable\n'
}
res=$(fn_test)
st=$?
[ "$st" -eq 42 ] || fail "function return code: want 42, got $st"
[ -z "$res" ] || fail "unreachable code executed: got [$res]"
echo PASS
exit 0
