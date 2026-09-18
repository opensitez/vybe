#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/disable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 8 % 2 == 0 )); then
  out=$(builtin echo -n "ok8")
else
  out=$(builtin echo -- "--ok8")
fi
if (( 8 % 2 == 0 )); then
  [ "$out" = "ok8" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok8" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
