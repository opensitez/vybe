#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_base_literals_octal_hex_and_decimal
# Leading-0 and 0x are interpreted as octal/hex in arithmetic contexts.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
value=$((010 + 0x10 + 10))
[ "$value" -eq 34 ] || fail "wanted 34, got $value"
echo PASS
exit 0
