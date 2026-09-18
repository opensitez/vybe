#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/final
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 20 % 2 == 0 )); then
  out=$(builtin echo -n "ok20")
else
  out=$(builtin echo -- "--ok20")
fi
if (( 20 % 2 == 0 )); then
  [ "$out" = "ok20" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok20" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
