#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_with_fallback_default_value
# Combining indirect expansion with default syntax ${!ptr:-default} falls back when target is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset phantom_target
ptr="phantom_target"
res="${!ptr:-backup_value}"
[ "$res" = "backup_value" ] || fail "fallback on unset target failed: got [$res]"
echo PASS
exit 0
