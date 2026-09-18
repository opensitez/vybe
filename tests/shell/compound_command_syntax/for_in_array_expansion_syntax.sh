#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/for_in_array_expansion_syntax
# The for name in "${arr[@]}" compound command iterates over each array element as an exact argument.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
items=("first item" "second item" "third")
concat=""
for elem in "${items[@]}"; do
    concat+="[${elem}]"
done
expected="[first item][second item][third]"
[ "$concat" = "$expected" ] || fail "array iteration: got [$concat]"
echo PASS
exit 0
