#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_lazy_evaluation_skips_command_sub
# If the variable is set and non-null, command substitutions inside default word are never executed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
valid_var="present"
sub_executed=0
res="${valid_var:-$(sub_executed=1; printf 'executed\n')}"
[ "$res" = "present" ] || fail "expansion mismatch: got [$res]"
[ "$sub_executed" -eq 0 ] || fail "default command substitution was eagerly executed"
echo PASS
exit 0
