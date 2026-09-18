#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_with_embedded_spaces_and_quotes
# Alternate words preserve internal spaces and quotes when the expansion is quoted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
opt="active"
res="${opt:+--option 'single quote' and \"double quote\"}"
expected="--option 'single quote' and \"double quote\""
[ "$res" = "$expected" ] || fail "quoted alternate corrupted: got [$res]"
echo PASS
exit 0
