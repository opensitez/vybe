#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/equal_in_double_brackets_is_pattern_match_not_assignment
# Inside [[ ... ]], '=' and '==' perform pattern or string equality matching, not variable assignment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample="abcdef"
[[ $sample = abc* ]] || fail "pattern match with = failed"
[[ $sample == abc* ]] || fail "pattern match with == failed"
echo PASS
exit 0
