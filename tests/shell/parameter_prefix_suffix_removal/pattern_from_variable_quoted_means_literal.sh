#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/pattern_from_variable_quoted_means_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=a.b.c
p='*.'
[ "${x##$p}" = c ] || fail "unquoted variable is a pattern: got [${x##$p}]"
[ "${x##"$p"}" = a.b.c ] || fail "quoted variable is literal text: got [${x##"$p"}]"
y='*.rest'
[ "${y#"$p"}" = rest ] || fail "literal *. removed: got [${y#"$p"}]"
echo PASS
exit 0
