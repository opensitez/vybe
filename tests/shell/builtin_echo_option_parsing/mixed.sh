#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/mixed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 17 % 2 == 0 )); then
  out=$(builtin echo -n "ok17")
else
  out=$(builtin echo -- "--ok17")
fi
if (( 17 % 2 == 0 )); then
  [ "$out" = "ok17" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok17" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
