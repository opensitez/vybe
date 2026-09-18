#!/usr/bin/env bash
# vybe-test: bash/test_builtin_forms/deprecated_a_o_and_parenthesis_grouping
# -a binds tighter than -o: "1 -eq 1 -o 1 -eq 2 -a 2 -eq 3" is
# true || (false && false), which is true; with -o tighter it would be false.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ 1 -eq 1 -a 2 -eq 2 ] || fail "-a true"
[ 1 -eq 1 -a 2 -eq 3 ] && fail "-a false"
[ 1 -eq 2 -o 2 -eq 2 ] || fail "-o true"
[ \( 1 -eq 2 -o 1 -eq 1 \) -a 2 -eq 2 ] || fail "parentheses group"
[ 1 -eq 1 -o 1 -eq 2 -a 2 -eq 3 ] || fail "-a binds tighter than -o"
[ \( 1 -eq 1 -o 1 -eq 2 \) -a 2 -eq 3 ] && fail "grouping the -o first makes it false"
echo PASS
exit 0
