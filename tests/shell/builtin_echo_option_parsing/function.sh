#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/function
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 2 % 2 == 0 )); then
  out=$(builtin echo -n "ok2")
else
  out=$(builtin echo -- "--ok2")
fi
if (( 2 % 2 == 0 )); then
  [ "$out" = "ok2" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok2" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
