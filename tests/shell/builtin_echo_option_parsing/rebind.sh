#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 16 % 2 == 0 )); then
  out=$(builtin echo -n "ok16")
else
  out=$(builtin echo -- "--ok16")
fi
if (( 16 % 2 == 0 )); then
  [ "$out" = "ok16" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok16" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
