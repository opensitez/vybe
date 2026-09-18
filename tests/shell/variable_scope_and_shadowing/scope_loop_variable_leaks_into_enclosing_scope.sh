#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_loop_variable_leaks_into_enclosing_scope
# In Bash, loops do not create scope; 'for item in ...' sets item in the enclosing environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset iter_item
for iter_item in 10 20 30; do
    :
done
[ "$iter_item" -eq 30 ] || fail "loop variable should leak into enclosing scope: got [$iter_item]"
echo PASS
exit 0
