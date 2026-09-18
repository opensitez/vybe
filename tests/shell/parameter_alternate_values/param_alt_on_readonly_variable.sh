#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_on_readonly_variable
# Alternate value expansions operate cleanly without error when applied to readonly variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly CONST_VAR="locked_val"
res="${CONST_VAR:+alternate_output}"
[ "$res" = "alternate_output" ] || fail "alternate on readonly variable failed: got [$res]"
[ "$CONST_VAR" = "locked_val" ] || fail "readonly variable altered"
echo PASS
exit 0
