#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/sign_and_surrounding_blanks_accepted
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ " 5" -eq 5 ] || fail "leading blank"
[ +5 -eq 5 ] || fail "explicit plus"
[ -5 -lt 0 ] || fail "negative"
[[ " 5" -eq 5 ]] || fail "leading blank in [["
echo PASS
exit 0
