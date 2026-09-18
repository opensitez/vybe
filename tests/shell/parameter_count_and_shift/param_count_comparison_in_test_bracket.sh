#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/param_count_comparison_in_test_bracket
# The parameter count $# can be verified using integer comparison operators in [[ ... ]] and [ ... ].
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "alpha" "beta"
[ "$#" -eq 2 ] || fail "[ \$# -eq 2 ] failed"
[ "$#" -lt 5 ] || fail "[ \$# -lt 5 ] failed"
[[ $# -ge 2 ]] || fail "[[ \$# -ge 2 ]] failed"
echo PASS
exit 0
