#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/locale_quoting_expands_parameters
# Locale translation $"..." behaves like double quotes and performs parameter expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
name="world"
val=$"hello $name"
[ "$val" = "hello world" ] || fail "locale quoting expansion: want 'hello world', got [$val]"
echo PASS
exit 0
