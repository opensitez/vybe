#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/blank_lines_ignored_between_commands
# Blank lines and lines containing only whitespace between commands are ignored by the parser.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }

step1="ok"


step2="done"

[ "$step1" = "ok" ] && [ "$step2" = "done" ] || fail "blank line handling failed"
echo PASS
exit 0
