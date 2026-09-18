#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_redirection_target
# The target word of a redirection undergoes tilde expansion. ~+ (the value of
# PWD) is used so the test stays inside its own temporary directory and does
# not have to reassign HOME.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp_dir=$(mktemp -d) || fail "mktemp failed"
trap 'rm -rf "$tmp_dir"' EXIT
cd "$tmp_dir" || fail "cd failed"
echo "payload" > ~+/test_output.txt
[ -f "$PWD/test_output.txt" ] || fail "redirection target ~+ did not expand to PWD"
content=$(<"$PWD/test_output.txt")
[ "$content" = "payload" ] || fail "file content: want [payload] got [$content]"
echo PASS
exit 0
