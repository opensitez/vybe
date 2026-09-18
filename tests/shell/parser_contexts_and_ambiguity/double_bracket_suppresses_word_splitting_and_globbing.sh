#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/double_bracket_suppresses_word_splitting_and_globbing
# Inside [[ ... ]], unquoted variables are not split into multiple words and glob characters are not expanded.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
spaced_var="hello world"
[[ -n $spaced_var ]] || fail "unquoted variable in [[ ]] should not split words"
[[ $spaced_var == "hello world" ]] || fail "unquoted variable equality in [[ ]]"
echo PASS
exit 0
