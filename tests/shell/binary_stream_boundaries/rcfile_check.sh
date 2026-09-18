#!/usr/bin/env bash
# vybe-test: bash/binary_stream_boundaries/rcfile_check
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=20
chunk_size=$((IDX % 5 + 1))
read -r -N "$chunk_size" chunk <<< "stream${IDX}"
(( ${#chunk} == chunk_size )) || fail "read -N did not get expected byte count"
if ! read -r -N 10 short <<< "abc"; then
  [[ $short == abc ]] || fail "short read should still collect available bytes"
else
  [[ $short == abc ]] || fail "short read expected short status"
fi
echo PASS
exit 0
