#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_ascii_string
# The ${#var} expansion returns the number of ASCII characters in the scalar variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample="hello_world"
[ "${#sample}" -eq 11 ] || fail "ASCII length: want 11, got ${#sample}"
echo PASS
exit 0
