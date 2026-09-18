#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_regex_capture/stream_boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=12
if [[ "item-${IDX}" =~ ^([a-z]+)-([0-9]+)$ ]]; then
  [[ ${BASH_REMATCH[1]} == item ]] || fail "capture group 1"
  [[ ${BASH_REMATCH[2]} == ${IDX} ]] || fail "capture group 2"
else
  fail "regex capture failed"
fi
echo PASS
exit 0
