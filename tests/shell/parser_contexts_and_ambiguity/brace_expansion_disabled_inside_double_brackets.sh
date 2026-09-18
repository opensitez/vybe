#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/brace_expansion_disabled_inside_double_brackets
# Brace expansion {a,b} is not performed inside [[ ... ]] expressions; braces remain literal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
item="{a,b}"
[[ $item == {a,b} ]] || fail "literal braces should match inside [[ ]]"
[[ "a" == {a,b} ]] && fail "brace expansion should NOT occur inside [[ ]]"
echo PASS
exit 0
