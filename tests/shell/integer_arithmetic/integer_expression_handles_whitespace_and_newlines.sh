#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_expression_handles_whitespace_and_newlines
# Newlines/tabs in arithmetic are tokenized as whitespace.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
value=$((1 + \
   2 + \
   (3 * 4) ))
[ "$value" -eq 15 ] || fail "wanted 15 got $value"
echo PASS
exit 0
