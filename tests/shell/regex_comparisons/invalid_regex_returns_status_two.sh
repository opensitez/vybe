#!/usr/bin/env bash
# vybe-test: bash/regex_comparisons/invalid_regex_returns_status_two
# 0 = match, 1 = no match, 2 = the regex itself is malformed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
re='('
msg=$( [[ a =~ $re ]] 2>&1 ); st=$?
[ "$st" -eq 2 ] || fail "want status 2 got $st"
[[ $msg == *"invalid regular expression"* ]] || fail "got [$msg]"
re='a{2'
[[ aa =~ $re ]] 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "unbalanced brace want 2 got $st"
echo PASS
exit 0
