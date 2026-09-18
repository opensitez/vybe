#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_ternary_only_takes_selected_branch
# Only the selected branch is evaluated, so the division-by-zero branch is safe when skipped.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 ? 11 : 2/0)) -eq 11 ] || fail "true branch got $((1 ? 11 : 2/0))"
[ $((0 ? 2/0 : 22)) -eq 22 ] || fail "false branch got $((0 ? 2/0 : 22))"
echo PASS
exit 0
