#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/arithmetic_inc_dec_in_for_loop_header
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
for ((i=0; i<4; ++i)); do count=$((count + i)); done
[ "$i" -eq 4 ] || fail "final i=$i"
[ "$count" -eq 6 ] || fail "sum=$count"
for ((j=6; j-- > 3; )); do :; done
after=$?
[ "$after" -eq 0 ] || fail "loop should finish successfully"
[ "$j" -eq 3 ] || fail "j should be 3 got $j"
echo PASS
exit 0
