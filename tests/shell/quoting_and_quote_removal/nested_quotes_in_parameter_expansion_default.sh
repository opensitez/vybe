#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/nested_quotes_in_parameter_expansion_default
# Quotes inside the default value of a parameter expansion ${var:-"..."} undergo proper quote removal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset unset_var
res="${unset_var:-"default value"}"
[ "$res" = "default value" ] || fail "nested quotes in default: want 'default value', got [$res]"
echo PASS
exit 0
