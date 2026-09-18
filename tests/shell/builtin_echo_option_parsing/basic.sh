#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 1 % 2 == 0 )); then
  out=$(builtin echo -n "ok1")
else
  out=$(builtin echo -- "--ok1")
fi
if (( 1 % 2 == 0 )); then
  [ "$out" = "ok1" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok1" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
