#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/unmatched_pattern_is_error_and_command_does_not_run
# The diagnostic is printed during expansion, before the command's own
# redirections exist, so it is captured on an enclosing group.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s failglob
msg=$( { echo ran *.zzz; } 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"no match: *.zzz"* ]] || fail "want no match diagnostic got [$msg]"
[[ $msg != *ran* ]] || fail "command must not run"
echo PASS
exit 0
