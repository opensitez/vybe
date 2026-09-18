#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/pattern_from_variable_quoted_means_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a*b.c'
p='*'
[ "${x//$p/_}" = '_' ] || fail "unquoted * pattern eats all: got [${x//$p/_}]"
[ "${x//"$p"/_}" = 'a_b.c' ] || fail "quoted * is literal: got [${x//"$p"/_}]"
echo PASS
exit 0
