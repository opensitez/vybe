#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_single_digit_positional_parameter
# The ${#1} through ${#9} expansions calculate string lengths of single-digit positional parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "short" "much_longer_string"
[ "${#1}" -eq 5 ] || fail "\${#1}: want 5, got ${#1}"
[ "${#2}" -eq 18 ] || fail "\${#2}: want 18, got ${#2}"
echo PASS
exit 0
