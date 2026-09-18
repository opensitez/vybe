#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_special_param_hash
# The ${##} expansion evaluates to the character length of the positional parameter count $#.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "arg"
[ "${##}" -eq 1 ] || fail "len of single-digit count: want 1, got ${##}"

set -- 1 2 3 4 5 6 7 8 9 10
[ "${##}" -eq 2 ] || fail "len of two-digit count: want 2, got ${##}"
echo PASS
exit 0
