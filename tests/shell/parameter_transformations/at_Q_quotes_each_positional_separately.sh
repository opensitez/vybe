#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/at_Q_quotes_each_positional_separately
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
set -- "a b" c
[ "$(count "${@@Q}")" = 2 ] || fail "\${@@Q} keeps one word per positional"
first() { echo "$1"; }
[ "$(first "${@@Q}")" = "'a b'" ] || fail "first quoted: got [$(first "${@@Q}")]"
[ "${*@Q}" = "'a b' 'c'" ] || fail "\${*@Q} joins: got [${*@Q}]"
echo PASS
exit 0
