#!/usr/bin/env bash
# vybe-test: bash/base64_command_interoperability/version_probe
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=18
base64(){ printf 'base64_%s' "$IDX"; }
out=$(base64)
[[ $out == "base64_${IDX}" ]] || fail "function name base64 mismatch"
unset -f base64
if (( IDX % 2 == 0 )); then
  if command -v base64 >/dev/null 2>&1; then
    :
  else
    :
  fi
fi
echo PASS
exit 0
