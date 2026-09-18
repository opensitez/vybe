#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/spaces_in_command_substitution_preserved_in_output
# Quoting a command substitution preserves its internal spaces and newlines verbatim.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="$(printf 'line 1   spaces\nline 2')"
expected="line 1   spaces"$'\n'"line 2"
[ "$out" = "$expected" ] || fail "command substitution whitespace preservation: got [$out]"
echo PASS
exit 0
