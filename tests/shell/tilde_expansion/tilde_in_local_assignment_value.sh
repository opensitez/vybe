#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_local_assignment_value
# The local builtin expands tildes in its assignment arguments inside a function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
test_func() {
    local my_path=~/local_dir
    [ "$my_path" = "$HOME/local_dir" ] || fail "local tilde failed: want [$HOME/local_dir], got [$my_path]"
}
test_func
echo PASS
exit 0
