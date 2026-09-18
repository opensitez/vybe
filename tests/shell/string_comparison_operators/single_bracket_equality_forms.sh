#!/usr/bin/env bash
# vybe-test: bash/string_comparison_operators/single_bracket_equality_forms
# In [ ] and test, = is the POSIX operator and == is a bash extension; both
# compare literally.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ abc = abc ] || fail "[ = ]"
[ abc == abc ] || fail "[ == ] extension"
test abc = abc || fail "test ="
[ abc != abd ] || fail "[ != ]"
[ abc = abd ] && fail "= false case"
[ abc != abc ] && fail "!= false case"
echo PASS
exit 0
