#!/usr/bin/env bash
# vybe-test: bash/nullglob_behavior/unmatched_pattern_as_command_name_runs_nothing
# When the command word itself vanishes the line is an empty command: status 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
./*.zzz 2>/dev/null; st=$?
[ "$st" -eq 127 ] || fail "without nullglob the literal name is not found, want 127 got $st"
shopt -s nullglob
./*.zzz; st=$?
[ "$st" -eq 0 ] || fail "with nullglob want 0 got $st"
echo PASS
exit 0
