#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/double_bracket_regex_match_syntax
# The [[ expr =~ regex ]] compound command syntax performs regular expression matching.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample="version-2.15.0"
[[ "$sample" =~ ^version-([0-9]+)\.([0-9]+) ]] || fail "regex match failed"
[ "${BASH_REMATCH[1]}" = "2" ] || fail "BASH_REMATCH[1]: want '2', got [${BASH_REMATCH[1]}]"
[ "${BASH_REMATCH[2]}" = "15" ] || fail "BASH_REMATCH[2]: want '15', got [${BASH_REMATCH[2]}]"
echo PASS
exit 0
