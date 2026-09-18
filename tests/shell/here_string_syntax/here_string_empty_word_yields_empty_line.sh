#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_empty_word_yields_empty_line
# An empty word <<< "" feeds an empty string followed by newline, succeeding on read with length 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r line <<< ""
st=$?
[ "$st" -eq 0 ] || fail "read on empty here-string: want status 0, got $st"
[ -z "$line" ] || fail "line should be empty, got [$line]"
echo PASS
exit 0
