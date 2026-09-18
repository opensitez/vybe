#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_length_expansion_hash_syntax
# The ${#1} expansion calculates the character length of positional parameter $1.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "hello_world" "12345"
[ "${#1}" -eq 11 ] || fail "length of \$1: want 11, got ${#1}"
[ "${#2}" -eq 5 ] || fail "length of \$2: want 5, got ${#2}"
echo PASS
exit 0
