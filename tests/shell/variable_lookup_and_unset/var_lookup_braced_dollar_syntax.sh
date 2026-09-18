#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_braced_dollar_syntax
# Braced ${var} syntax unambiguously delimits the variable name from adjacent text.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
stem="branch"
combined="${stem}_suffix"
[ "$combined" = "branch_suffix" ] || fail "braced lookup failed: got [$combined]"
echo PASS
exit 0
