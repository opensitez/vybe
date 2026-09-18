#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_special_param_underscore
# The ${#_} expansion returns the character length of the $_ parameter (last argument).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: "1234567"
[ "${#_}" -eq 7 ] || fail "len of \$_: want 7, got ${#_}"
echo PASS
exit 0
