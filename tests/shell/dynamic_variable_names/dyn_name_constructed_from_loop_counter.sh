#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_name_constructed_from_loop_counter
# Variable names can be generated in a loop and dynamically read via indirect expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
for i in 1 2 3; do
    printf -v "item_$i" "data_%d" $(( i * 10 ))
done
for i in 1 2 3; do
    ptr="item_$i"
    expected="data_$(( i * 10 ))"
    [ "${!ptr}" = "$expected" ] || fail "loop-generated variable $ptr: want [$expected], got [${!ptr}]"
done
echo PASS
exit 0
