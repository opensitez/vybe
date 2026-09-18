#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_Q_quotes_string_for_safe_reuse
# The ${var@Q} parameter transformation quotes the variable's value safely for reuse as shell input.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
raw_input="arbitrary string with 'single' and \"double\" quotes"
quoted="${raw_input@Q}"
# Evaluating the quoted string reproduces the exact raw input
eval "reproduced=$quoted"
[ "$reproduced" = "$raw_input" ] || fail "eval of \${var@Q} corrupted value: got [$reproduced]"
echo PASS
exit 0
