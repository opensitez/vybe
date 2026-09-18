#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/set_double_dash_resets_count_after_shift
# Re-invoking 'set --' completely resets the positional parameter list and count after shifting.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "old1" "old2" "old3"
shift 2
[ "$#" -eq 1 ] || fail "intermediate shift failed"
set -- "fresh_a" "fresh_b" "fresh_c" "fresh_d"
[ "$#" -eq 4 ] || fail "parameter count after reset: want 4, got $#"
[ "$1" = "fresh_a" ] || fail "param 1: want 'fresh_a', got [$1]"
echo PASS
exit 0
