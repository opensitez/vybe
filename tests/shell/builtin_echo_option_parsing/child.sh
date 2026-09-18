#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/child
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 11 % 2 == 0 )); then
  out=$(builtin echo -n "ok11")
else
  out=$(builtin echo -- "--ok11")
fi
if (( 11 % 2 == 0 )); then
  [ "$out" = "ok11" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok11" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
