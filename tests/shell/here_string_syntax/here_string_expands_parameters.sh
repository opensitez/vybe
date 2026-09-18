#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_expands_parameters
# The word argument to <<< undergoes parameter expansion before being presented to standard input.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="expanded_value"
read -r result <<< "prefix_${var}"
[ "$result" = "prefix_expanded_value" ] || fail "parameter expansion in here-string: got [$result]"
echo PASS
exit 0
