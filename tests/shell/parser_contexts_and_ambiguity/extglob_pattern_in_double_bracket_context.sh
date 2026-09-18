#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/extglob_pattern_in_double_bracket_context
# Extended glob patterns like +(pattern) are parsed in [[ ... ]] without pathname expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
word="aaa"
[[ $word == +(a) ]] || fail "extglob +(a) match failed"
[[ "aab" == +(a) ]] && fail "extglob +(a) should not match 'aab'"
echo PASS
exit 0
