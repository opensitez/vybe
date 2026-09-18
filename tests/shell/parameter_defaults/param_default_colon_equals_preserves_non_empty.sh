#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_equals_preserves_non_empty
# The ${var:=default} syntax preserves var's value without reassigning when var is non-empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pre_set="initial_payload"
res="${pre_set:=fallback_payload}"
[ "$res" = "initial_payload" ] || fail "expansion mismatch: got [$res]"
[ "$pre_set" = "initial_payload" ] || fail "non-empty variable was overwritten by :=: got [$pre_set]"
echo PASS
exit 0
