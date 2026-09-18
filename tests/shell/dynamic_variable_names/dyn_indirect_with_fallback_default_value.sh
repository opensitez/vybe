#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_indirect_with_fallback_default_value
# Combining indirect expansion with default fallback ${!ptr:-default} falls back when target is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset undefined_target
ptr="undefined_target"
val="${!ptr:-fallback_val}"
[ "$val" = "fallback_val" ] || fail "indirect fallback on unset target failed: got [$val]"

defined_target="active_val"
ptr="defined_target"
val2="${!ptr:-fallback_val}"
[ "$val2" = "active_val" ] || fail "indirect fallback on defined target failed: got [$val2]"
echo PASS
exit 0
