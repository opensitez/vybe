#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/semicolon_binds_looser_than_and_or
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(false && echo a; echo b)
[ "$out" = b ] || fail "got [$out]"
echo PASS
exit 0
