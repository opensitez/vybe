#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_positional_parameter_alternate
# The alternate value expansion applies to positional parameters $1, $2, etc.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "arg1"
res1="${1:+has_arg1}"
res2="${2:+has_arg2}"
[ "$res1" = "has_arg1" ] || fail "\$1 alternate failed: got [$res1]"
[ -z "$res2" ] || fail "\$2 alternate should be empty: got [$res2]"
echo PASS
exit 0
