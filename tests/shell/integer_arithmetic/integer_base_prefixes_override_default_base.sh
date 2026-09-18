#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_base_prefixes_override_default_base
# Base prefixes force parsing independent of variable style.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
value=$((10#12 + 2#101 + 8#10 + 16#f))
[ "$value" -eq 40 ] || fail "wanted 40, got $value"
echo PASS
exit 0
