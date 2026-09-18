#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_numeric_sequence_zero_padded_width
# Using leading zeros like {01..05} forces the expanded numbers to be zero-padded to equal width.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {01..05}
[ "$#" -eq 5 ] || fail "padded count: want 5, got $#"
[ "$*" = "01 02 03 04 05" ] || fail "zero-padded sequence mismatch: got [$*]"

set -- {08..12}
[ "$*" = "08 09 10 11 12" ] || fail "multi-digit padding mismatch: got [$*]"
echo PASS
exit 0
