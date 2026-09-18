#!/usr/bin/env bash
# vybe-test: bash/process_substitution/source_from_process_substitution
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source <(echo 'v=fromproc; f() { echo fn; }')
[ "$v" = fromproc ] || fail "variable: got [$v]"
[ "$(f)" = fn ] || fail "function not defined"
echo PASS
exit 0
