#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/option_test_o_in_single_bracket
# -o optname is a shell-state test, not a file test; it works in [ ] as well.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set +f
[ -o noglob ] && fail "noglob off"
set -f
[ -o noglob ] || fail "noglob on in [ ]"
test -o noglob || fail "noglob on in test"
[ -o nosuchoption ] && fail "unknown option name must be false"
echo PASS
exit 0
