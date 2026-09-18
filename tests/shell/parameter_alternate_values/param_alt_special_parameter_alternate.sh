#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_special_parameter_alternate
# The alternate value expansion applies to special parameters such as $# and $?.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "sample"
res_hash="${#:+has_args}"
[ "$res_hash" = "has_args" ] || fail "\$# alternate failed: got [$res_hash]"

true
res_q="${?:+has_status}"
[ "$res_q" = "has_status" ] || fail "\$? alternate failed: got [$res_q]"
echo PASS
exit 0
