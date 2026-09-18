#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/empty_replacement_deletes_matches
# Both ${x/pat/} and the shorter ${x/pat} delete the match.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=banana
[ "${x/a/}" = bnana ] || fail "got [${x/a/}]"
[ "${x/a}" = bnana ] || fail "got [${x/a}]"
[ "${x//a}" = bnn ] || fail "got [${x//a}]"
echo PASS
exit 0
