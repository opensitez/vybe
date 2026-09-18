#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_two_digit_positional_parameter
# The ${#10} expansion calculates the string character length of the 10th positional parameter.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- a b c d e f g h i "tenth_argument_content"
[ "${#10}" -eq 22 ] || fail "\${#10}: want 22, got ${#10}"
echo PASS
exit 0
