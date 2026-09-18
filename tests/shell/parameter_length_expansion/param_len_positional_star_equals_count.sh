#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_positional_star_equals_count
# The ${#*} expansion evaluates to the decimal positional parameter count, equivalent to $#.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "a" "b" "c" "d"
[ "${#*}" -eq 4 ] || fail "\${#*}: want 4, got ${#*}"
[ "${#*}" -eq "$#" ] || fail "\${#*} does not match \$#"
echo PASS
exit 0
