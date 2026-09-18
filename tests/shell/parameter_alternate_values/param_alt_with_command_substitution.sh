#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_with_command_substitution
# The alternate word evaluates command substitutions $( ... ) when condition is met.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
service_active="true"
res="${service_active:+$(printf 'status:online\n')}"
[ "$res" = "status:online" ] || fail "command substitution in alternate failed: got [$res]"
echo PASS
exit 0
