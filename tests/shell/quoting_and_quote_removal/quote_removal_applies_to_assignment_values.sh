#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quote_removal_applies_to_assignment_values
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="a"'b'c\d
[ "$x" = abcd ] || fail "want [abcd] got [$x]"
[ "${#x}" -eq 4 ] || fail "length want 4 got ${#x}"
echo PASS
exit 0
