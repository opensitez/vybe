#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 18 % 2 == 0 )); then
  out=$(builtin echo -n "ok18")
else
  out=$(builtin echo -- "--ok18")
fi
if (( 18 % 2 == 0 )); then
  [ "$out" = "ok18" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok18" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
