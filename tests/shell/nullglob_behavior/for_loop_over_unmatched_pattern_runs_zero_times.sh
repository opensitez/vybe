#!/usr/bin/env bash
# vybe-test: bash/nullglob_behavior/for_loop_over_unmatched_pattern_runs_zero_times
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
n=0; for f in *.zzz; do n=$((n+1)); done
[ "$n" -eq 1 ] || fail "without nullglob the literal pattern is iterated once, got $n"
shopt -s nullglob
n=0; for f in *.zzz; do n=$((n+1)); done
[ "$n" -eq 0 ] || fail "with nullglob want 0 got $n"
echo PASS
exit 0
