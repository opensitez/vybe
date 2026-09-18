#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/error_abandons_the_rest_of_the_command_line
# Like other expansion errors, a failglob failure discards everything else in
# the same top-level command list: a command after ; on that line never runs.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s failglob
unset touched
{ echo *.zzz; } 2>/dev/null; touched=yes
[ -z "${touched+set}" ] || fail "command after ; on the failing line must not run"
next_line=ran
[ "$next_line" = ran ] || fail "unreachable"
echo PASS
exit 0
