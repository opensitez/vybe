#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_whitespace_preservation
# Quoted scalar assignments preserve leading, trailing, and consecutive internal whitespace characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
spaced="   leading,   multiple-internal, and trailing   "
[ "$spaced" = "   leading,   multiple-internal, and trailing   " ] || fail "whitespace preservation failed"
[ "${#spaced}" -eq 48 ] || fail "length mismatch: want 48, got ${#spaced}"
echo PASS
exit 0
