#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/function_names_are_case_sensitive
# Function identifiers differing in letter case define distinct shell functions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fn_test() { printf 'lower\n'; }
Fn_Test() { printf 'title\n'; }
[ "$(fn_test)" = "lower" ] || fail "fn_test: want 'lower', got [$(fn_test)]"
[ "$(Fn_Test)" = "title" ] || fail "Fn_Test: want 'title', got [$(Fn_Test)]"
echo PASS
exit 0
