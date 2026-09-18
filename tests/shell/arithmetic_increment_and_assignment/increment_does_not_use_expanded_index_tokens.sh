#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/increment_does_not_use_expanded_index_tokens
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
# In (( )), variable names are parsed as names, then their values are arithmetic.
idx='1'
a=(9 8 7)
((idx++))
[ "$idx" -eq 2 ] || fail "idx should become 2"
[ "${a[idx]}" = 7 ] || fail "a[2] after idx increment"
[ "$idx" -eq 2 ] || fail "idx check"
echo PASS
exit 0
