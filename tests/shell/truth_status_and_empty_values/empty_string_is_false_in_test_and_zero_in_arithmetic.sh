#!/usr/bin/env bash
# vybe-test: bash/truth_status_and_empty_values/empty_string_is_false_in_test_and_zero_in_arithmetic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
e=
[ "$e" ] && fail "[ \"\" ] must be false"
[[ $e ]] && fail "[[ \$e ]] must be false"
(( e )) && fail "(( e )) with empty e is 0: false"
[ $((e + 1)) -eq 1 ] || fail "empty is 0 in arithmetic"
echo PASS
exit 0
