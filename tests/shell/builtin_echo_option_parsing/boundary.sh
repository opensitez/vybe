#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 13 % 2 == 0 )); then
  out=$(builtin echo -n "ok13")
else
  out=$(builtin echo -- "--ok13")
fi
if (( 13 % 2 == 0 )); then
  [ "$out" = "ok13" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok13" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
