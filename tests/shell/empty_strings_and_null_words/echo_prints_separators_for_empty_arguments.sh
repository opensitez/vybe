#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/echo_prints_separators_for_empty_arguments
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo "" ""; echo x)
[ "$out" = $' \nx' ] || fail "two empty args give one space, got [$out]"
out=$(echo ""; echo x)
[ "$out" = $'\nx' ] || fail "one empty arg gives an empty line, got [$out]"
echo PASS
exit 0
