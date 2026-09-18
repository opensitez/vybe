#!/usr/bin/env bash
# vybe-test: bash/command_substitution/nesting_with_independent_quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo $(echo in))
[ "$out" = in ] || fail "plain nesting: got [$out]"
out="$(echo "$(echo "a  b")")"
[ "$out" = "a  b" ] || fail "quoted nesting keeps spaces: got [$out]"
echo PASS
exit 0
