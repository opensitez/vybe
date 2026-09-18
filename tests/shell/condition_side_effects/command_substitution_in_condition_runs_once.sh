#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/command_substitution_in_condition_runs_once
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT
counter() { echo . >> "$tmp/calls"; echo x; }
if [ "$(counter)" = x ]; then :; fi
n=0; while read -r _; do n=$((n+1)); done < "$tmp/calls"
[ "$n" -eq 1 ] || fail "want 1 call got $n"
echo PASS
exit 0
