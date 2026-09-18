#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/function_body_must_be_compound_command
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'f() echo hi' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "simple command as function body must be a syntax error"
eval 'g() ( echo sub )'; st=$?
[ "$st" -eq 0 ] || fail "subshell body is a valid compound command"
[ "$(g)" = sub ] || fail "g should print sub"
echo PASS
exit 0
