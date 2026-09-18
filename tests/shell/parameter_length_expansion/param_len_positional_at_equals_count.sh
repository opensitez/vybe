#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_positional_at_equals_count
# The ${#@} expansion evaluates to the decimal positional parameter count, equivalent to $#.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "item1" "item2"
[ "${#@}" -eq 2 ] || fail "\${#@}: want 2, got ${#@}"
[ "${#@}" -eq "$#" ] || fail "\${#@} does not match \$#"
echo PASS
exit 0
