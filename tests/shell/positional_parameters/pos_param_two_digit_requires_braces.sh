#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_two_digit_requires_braces
# Accessing positional parameter 10+ requires braces ${10}; unbraced $10 expands to $1 followed by literal 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- a b c d e f g h i "tenth_val"
unbraced="$10"
braced="${10}"
[ "$unbraced" = "a0" ] || fail "\$10 should expand to \$1 followed by '0': got [$unbraced]"
[ "$braced" = "tenth_val" ] || fail "\${10} should expand to tenth parameter: got [$braced]"
echo PASS
exit 0
