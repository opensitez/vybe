#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_extglob_optional_parts/missing_pattern
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=5
shopt -s extglob
if (( IDX % 2 == 0 )); then
  [[ "foo" == foo?(bar) ]] || fail "optional group missing base form"
  [[ "foobar" == foo?(bar) ]] || fail "optional group should accept suffix"
else
  [[ "foobarbaz" == foo?(bar) ]] && fail "optional group should not consume extra text"
fi
shopt -u extglob
echo PASS
exit 0
