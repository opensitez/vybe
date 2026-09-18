#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/set_f_disables_pathname_expansion
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a.txt
set -f
[ "$(echo *.txt)" = '*.txt' ] || fail "noglob: got [$(echo *.txt)]"
[[ $- == *f* ]] || fail "f flag must show in \$-"
set +f
[ "$(echo *.txt)" = a.txt ] || fail "re-enabled: got [$(echo *.txt)]"
echo PASS
exit 0
