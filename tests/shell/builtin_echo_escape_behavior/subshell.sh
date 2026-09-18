#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_escape_behavior/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 3 % 2 == 0 )); then
  out="$(builtin echo -e "a\tb")"
  expected=$'a\tb'
else
  out="$(builtin echo -e "a\nb")"
  expected=$'a\nb'
fi
[ "$out" = "$expected" ] || fail "echo escape mismatch"
echo PASS
exit 0
