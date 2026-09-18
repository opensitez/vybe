#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/pattern_from_variable_follows_the_same_quoting_rule
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc; p='a*'
[[ $x == $p ]] || fail "unquoted \$p is a pattern"
[[ $x == "$p" ]] && fail "quoted \$p is literal"
[[ $x == ${p:0:1}* ]] || fail "mixed: expansion plus literal glob"
echo PASS
exit 0
