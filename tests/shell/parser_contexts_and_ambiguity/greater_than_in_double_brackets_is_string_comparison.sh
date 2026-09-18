#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/greater_than_in_double_brackets_is_string_comparison
# Inside [[ ... ]], the '>' and '<' tokens are string lexicographical comparison operators, not file redirections.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ "banana" > "apple" ]] || fail "'banana' should be > 'apple'"
[[ "apple" < "banana" ]] || fail "'apple' should be < 'banana'"
echo PASS
exit 0
