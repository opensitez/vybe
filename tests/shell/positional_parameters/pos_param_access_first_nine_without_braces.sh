#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_access_first_nine_without_braces
# Positional parameters $1 through $9 can be accessed directly without curly braces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "one" "two" "three" "four" "five" "six" "seven" "eight" "nine"
[ "$1" = "one" ] || fail "\$1 mismatch: got [$1]"
[ "$5" = "five" ] || fail "\$5 mismatch: got [$5]"
[ "$9" = "nine" ] || fail "\$9 mismatch: got [$9]"
echo PASS
exit 0
