#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_double_quoted_with_spaces
# Double quotes preserve internal spaces within the supplied here-string argument.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IFS= read -r result <<< "  spaced   content  "
[ "$result" = "  spaced   content  " ] || fail "spaces in here-string altered: got [$result]"
echo PASS
exit 0
