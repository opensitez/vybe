#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/newline_allowed_after_pipe
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo abc |
{ read -r v; echo "got:$v"; })
[ "$out" = got:abc ] || fail "got [$out]"
echo PASS
exit 0
