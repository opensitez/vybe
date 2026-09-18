#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 7 % 2 == 0 )); then
  out=$(builtin echo -n "ok7")
else
  out=$(builtin echo -- "--ok7")
fi
if (( 7 % 2 == 0 )); then
  [ "$out" = "ok7" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok7" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
