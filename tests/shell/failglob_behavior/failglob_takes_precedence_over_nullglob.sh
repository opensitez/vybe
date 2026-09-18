#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/failglob_takes_precedence_over_nullglob
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s nullglob failglob
msg=$( { echo *.zzz; } 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"no match"* ]] || fail "got [$msg]"
echo PASS
exit 0
