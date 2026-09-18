#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/double_slash_replaces_every_match
# Matches are found left to right and replaced text is never rescanned.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=banana
[ "${x//a/A}" = bAnAnA ] || fail "got [${x//a/A}]"
y=aaa
[ "${y//a/aa}" = aaaaaa ] || fail "replacement must not be rescanned: got [${y//a/aa}]"
echo PASS
exit 0
