#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_division_by_zero_in_expansion_fails
# Division by zero errors in arithmetic expansion and does not silently coerce.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'echo $((1 / 0))' 2>&1); st=$?
[ "$st" -eq 1 ] || fail "want status 1 from expansion, got $st"
[[ $msg == *"division by zero"* ]] || fail "unexpected message: [$msg]"
echo PASS
exit 0
