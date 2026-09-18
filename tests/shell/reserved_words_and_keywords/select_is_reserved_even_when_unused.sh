#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/select_is_reserved_even_when_unused
# select needs a list; using it like a command is a syntax error, not a lookup.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'select' 2>&1); st=$?
[ "$st" -ne 0 ] || fail "bare select must fail"
[[ $msg == *"syntax error"* ]] || fail "want syntax error got [$msg]"
echo PASS
exit 0
