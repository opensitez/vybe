#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_with_arithmetic_expansion
# The default word expands arithmetic expressions when evaluated for unset or null parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset port_num
res="${port_num:-$(( 8000 + 80 ))}"
[ "$res" -eq 8080 ] || fail "arithmetic default failed: got [$res]"
echo PASS
exit 0
