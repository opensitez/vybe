#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/missing_closing_keyword_is_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'if true; then echo a' 2>/dev/null; a=$?
eval 'while true; do break' 2>/dev/null; b=$?
eval 'case x in x) echo a;;' 2>/dev/null; c=$?
eval 'f() { echo a;' 2>/dev/null; d=$?
[ "$a" -ne 0 ] || fail "missing fi"
[ "$b" -ne 0 ] || fail "missing done"
[ "$c" -ne 0 ] || fail "missing esac"
[ "$d" -ne 0 ] || fail "missing }"
echo PASS
exit 0
