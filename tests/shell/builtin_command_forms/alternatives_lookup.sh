#!/usr/bin/env bash
# vybe-test: bash/builtin_command_forms/alternatives_lookup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=2
command printf '%s' "form_${IDX}" >/dev/null || fail "command builtin failed"
builtin printf '%s' "form_${IDX}" >/dev/null || fail "builtin form failed"
if (( IDX % 2 == 0 )); then
  command -V printf >/dev/null 2>&1 || fail "command -V should return a type"
fi
if (( IDX % 3 == 0 )); then
  builtin -p printf >/dev/null 2>&1 || :
fi
echo PASS
exit 0
