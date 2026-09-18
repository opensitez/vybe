#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/error_does_not_terminate_the_script
# A failglob error aborts the current command line; a non-interactive shell
# continues with the next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s failglob
{ echo *.zzz; } 2>/dev/null
st=$?
[ "$st" -eq 1 ] || fail "want status 1 got [$st]"
after=reached
[ "$after" = reached ] || fail "unreachable"
echo PASS
exit 0
