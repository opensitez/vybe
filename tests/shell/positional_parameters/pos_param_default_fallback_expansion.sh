#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_default_fallback_expansion
# The ${1:-fallback} syntax provides default value when positional parameter is unset or empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --
val="${1:-default_value}"
[ "$val" = "default_value" ] || fail "fallback failed: got [$val]"

set -- "provided_value"
val2="${1:-default_value}"
[ "$val2" = "provided_value" ] || fail "present parameter replaced by default: got [$val2]"
echo PASS
exit 0
