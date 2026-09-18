#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/keywords_are_case_sensitive
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'IF true; then :; fi' 2>&1); st=$?
[ "$st" -ne 0 ] || fail "IF must not be a keyword"
[[ $msg == *"syntax error"* ]] || fail "want syntax error near then, got [$msg]"
t=$(type -t IF); [ -z "$t" ] || fail "type IF should be nothing, got [$t]"
echo PASS
exit 0
