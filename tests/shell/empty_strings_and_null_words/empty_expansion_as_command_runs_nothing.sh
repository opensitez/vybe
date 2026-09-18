#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/empty_expansion_as_command_runs_nothing
# A command line whose only word expands to nothing is an empty command
# (status 0); an explicit null word "" is a command named "" and is not found.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=
$x; st=$?
[ "$st" -eq 0 ] || fail "empty expansion: want 0 got $st"
"" 2>/dev/null; st=$?
[ "$st" -eq 127 ] || fail "null word: want 127 got $st"
echo PASS
exit 0
