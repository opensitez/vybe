#!/usr/bin/env bash
# vybe-test: bash/bash_version_features/alias_override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=9
(( BASH_VERSINFO[0] >= 1 )) || fail "major version invalid"
(( BASH_VERSINFO[1] >= 0 )) || fail "minor version invalid"
[[ -n "$BASH_VERSION" ]] || fail "BASH_VERSION missing"
if (( IDX % 2 == 0 )); then
  [[ $BASH_VERSION == *"${BASH_VERSINFO[0]}.${BASH_VERSINFO[1]}"* ]] || fail "version fields should align"
fi
echo PASS
exit 0
