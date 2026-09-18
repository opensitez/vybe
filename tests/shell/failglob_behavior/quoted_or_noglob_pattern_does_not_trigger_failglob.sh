#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/quoted_or_noglob_pattern_does_not_trigger_failglob
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s failglob
out=$(echo '*.zzz' 2>&1); st=$?
[ "$st" -eq 0 ] && [ "$out" = '*.zzz' ] || fail "quoted: st=$st got [$out]"
set -f
out=$(echo *.zzz 2>&1); st=$?
[ "$st" -eq 0 ] && [ "$out" = '*.zzz' ] || fail "noglob: st=$st got [$out]"
echo PASS
exit 0
