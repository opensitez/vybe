#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_regex_capture_groups
# Parentheses in a regex pattern inside [[ ... =~ ... ]] define subpattern capture groups into BASH_REMATCH.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample="server:8080"
[[ "$sample" =~ ^([a-z]+):([0-9]+)$ ]] || fail "regex capture match failed"
[ "${BASH_REMATCH[1]}" = "server" ] || fail "group 1: want 'server', got [${BASH_REMATCH[1]}]"
[ "${BASH_REMATCH[2]}" = "8080" ] || fail "group 2: want '8080', got [${BASH_REMATCH[2]}]"
echo PASS
exit 0
