#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/param_count_used_in_arithmetic_eval_syntax
# The parameter count $# can be used directly in arithmetic evaluations (( ... )) and $(( ... )).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- 10 20 30 40
(( doubled = $# * 2 ))
[ "$doubled" -eq 8 ] || fail "arithmetic evaluation on \$# failed: want 8, got $doubled"
sum=$(( $# + 100 ))
[ "$sum" -eq 104 ] || fail "\$(( \$# + 100 )): want 104, got $sum"
echo PASS
exit 0
