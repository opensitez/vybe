#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/cleanup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 15 % 2 == 0 )); then
  out=$(builtin echo -n "ok15")
else
  out=$(builtin echo -- "--ok15")
fi
if (( 15 % 2 == 0 )); then
  [ "$out" = "ok15" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok15" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
