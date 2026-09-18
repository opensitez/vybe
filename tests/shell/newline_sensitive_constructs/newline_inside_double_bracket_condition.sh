#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_inside_double_bracket_condition
# Newlines within [[ ... ]] expressions separate compound tests and operators without error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 1 -eq 1
   && "abc" == "abc"
   || 3 -eq 4 ]] || fail "multiline [[ ... ]] condition failed"
echo PASS
exit 0
