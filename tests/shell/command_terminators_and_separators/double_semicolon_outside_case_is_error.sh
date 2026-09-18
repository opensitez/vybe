#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/double_semicolon_outside_case_is_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'echo a;; echo b' 2>&1); st=$?
[ "$st" -ne 0 ] || fail ";; must be a syntax error"
[[ $msg == *"syntax error"* ]] || fail "got [$msg]"
echo PASS
exit 0
