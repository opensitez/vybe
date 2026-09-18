#!/usr/bin/env bash
# vybe-test: bash/string_comparison_operators/double_bracket_uses_locale_collation_single_uses_strcmp
# In the C locale upper case sorts before lower case in both forms; under a
# natural-language locale [[ ]] follows the locale's collation while [ ]
# keeps byte order.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export LC_ALL=C
[[ B < a ]] || fail "C locale [[ B < a ]] must be true"
[ B \< a ] || fail "C locale [ B < a ] must be true"
export LC_ALL=en_US.UTF-8
[[ B < a ]] && fail "en_US [[ B < a ]] must be false (a collates before B)"
[ B \< a ] || fail "[ ] must still use byte order under en_US"
echo PASS
exit 0
