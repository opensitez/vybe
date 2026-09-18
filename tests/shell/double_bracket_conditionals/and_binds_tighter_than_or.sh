#!/usr/bin/env bash
# vybe-test: bash/double_bracket_conditionals/and_binds_tighter_than_or
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ a == b || a == a && b == c ]] && fail "must parse as a==b || (a==a && b==c) which is false"
[[ ( a == b || a == a ) && b == c ]] && fail "grouped form is also false"
[[ a == a && b == c || c == c ]] || fail "(false) || true is true"
echo PASS
exit 0
