#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_as_parameter_length_operator
# The '#' in ${#var} is the parameter length operator, not a comment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
str="1234567"
len=${#str}
[ "$len" -eq 7 ] || fail "param length: want 7, got $len"
echo PASS
exit 0
