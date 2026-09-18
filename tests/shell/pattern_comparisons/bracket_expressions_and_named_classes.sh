#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/bracket_expressions_and_named_classes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 5 == [[:digit:]] ]] || fail "[[:digit:]]"
[[ b == [a-c] ]] || fail "range"
[[ d == [!a-c] ]] || fail "negated range"
[[ d == [^a-c] ]] || fail "caret negation"
[[ ab == [[:alpha:]][[:alpha:]] ]] || fail "two classes"
[[ 5 == [[:alpha:]] ]] && fail "digit is not alpha"
echo PASS
exit 0
