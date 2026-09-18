#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/parent
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 10 % 2 == 0 )); then
  out=$(builtin echo -n "ok10")
else
  out=$(builtin echo -- "--ok10")
fi
if (( 10 % 2 == 0 )); then
  [ "$out" = "ok10" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok10" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
