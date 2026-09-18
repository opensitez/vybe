#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_with_embedded_spaces_and_quotes
# Default words can contain quoted spaces and literals that are preserved when the expansion is quoted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset unassigned
res="${unassigned:-'literal single' and \"double quote\"}"
expected="'literal single' and \"double quote\""
[ "$res" = "$expected" ] || fail "quoted default corrupted: got [$res]"
echo PASS
exit 0
