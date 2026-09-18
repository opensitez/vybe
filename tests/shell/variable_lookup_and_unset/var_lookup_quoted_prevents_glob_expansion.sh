#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_quoted_prevents_glob_expansion
# Referencing "$var" where the value contains glob characters prevents pathname expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pattern="*"
pass_arg() { printf '%s\n' "$1"; }
result=$(pass_arg "$pattern")
[ "$result" = "*" ] || fail "quoted glob expanded into directory files: got [$result]"
echo PASS
exit 0
