#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_string_with_embedded_newlines
# The ${#var} expansion counts newline characters $'\n' accurately as individual characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
nl_str="a"$'\n'"b"$'\n'"c"
[ "${#nl_str}" -eq 5 ] || fail "newline string length: want 5, got ${#nl_str}"
echo PASS
exit 0
