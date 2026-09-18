#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/redirection_target_single_match_ok_multiple_ambiguous
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > only.txt; : > a.log; : > b.log
echo payload > *.txt; st=$?
[ "$st" -eq 0 ] || fail "single match must work, st=$st"
[ "$(<only.txt)" = payload ] || fail "content not written to the matched file"
msg=$( { echo payload > *.log; } 2>&1 ); st=$?
[ "$st" -ne 0 ] || fail "multiple matches must fail"
[[ $msg == *"ambiguous redirect"* ]] || fail "got [$msg]"
echo PASS
exit 0
