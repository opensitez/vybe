#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/ere_alternation_repetition_and_classes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ b =~ ^(a|b)$ ]] || fail "alternation"
[[ c =~ ^(a|b)$ ]] && fail "alternation non-member"
[[ aa =~ ^a{2}$ ]] || fail "exact repetition"
[[ aaa =~ ^a{2}$ ]] && fail "exact repetition too many"
[[ abc =~ ^[[:alpha:]]+$ ]] || fail "class with +"
[[ ab1 =~ ^[[:alpha:]]+$ ]] && fail "digit breaks alpha class"
[[ ac =~ ^ab*c$ ]] || fail "star zero times"
echo PASS
exit 0
