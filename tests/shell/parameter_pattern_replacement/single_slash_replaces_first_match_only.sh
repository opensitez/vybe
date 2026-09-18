#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/single_slash_replaces_first_match_only
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=banana
[ "${x/a/A}" = bAnana ] || fail "got [${x/a/A}]"
[ "${x/na/-}" = ba-na ] || fail "got [${x/na/-}]"
echo PASS
exit 0
