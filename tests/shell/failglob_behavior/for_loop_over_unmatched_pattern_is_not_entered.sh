#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/for_loop_over_unmatched_pattern_is_not_entered
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s failglob
n=0
for f in *.zzz; do n=$((n+1)); done 2>/dev/null
st=$?
[ "$n" -eq 0 ] || fail "body ran $n times"
[ "$st" -eq 1 ] || fail "loop status want 1 got $st"
echo PASS
exit 0
