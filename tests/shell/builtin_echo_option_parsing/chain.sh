#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 5 % 2 == 0 )); then
  out=$(builtin echo -n "ok5")
else
  out=$(builtin echo -- "--ok5")
fi
if (( 5 % 2 == 0 )); then
  [ "$out" = "ok5" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok5" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
