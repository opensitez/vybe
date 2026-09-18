#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_with_embedded_newlines
# Assigning a string with embedded newlines preserves line breaks accurately across expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline="row1
row2
row3"
line_count=0
while read -r line; do
    line_count=$(( line_count + 1 ))
done <<< "$multiline"
[ "$line_count" -eq 3 ] || fail "multiline assignment line count: want 3, got $line_count"
echo PASS
exit 0
