#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 19 % 2 == 0 )); then
  out=$(builtin echo -n "ok19")
else
  out=$(builtin echo -- "--ok19")
fi
if (( 19 % 2 == 0 )); then
  [ "$out" = "ok19" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok19" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
