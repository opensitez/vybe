#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_chained_two_variables
# Chaining alternate value expansions enables concise logical-AND conditional evaluation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var1="ready"
var2="configured"
res="${var1:+${var2:+both_are_set}}"
[ "$res" = "both_are_set" ] || fail "chained alternate failed when both set: got [$res]"

unset var2
res_partial="${var1:+${var2:+both_are_set}}"
[ -z "$res_partial" ] || fail "chained alternate should be empty when var2 unset: got [$res_partial]"
echo PASS
exit 0
