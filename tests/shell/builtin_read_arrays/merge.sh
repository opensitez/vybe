#!/usr/bin/env bash
# vybe-test: bash/builtin_read_arrays/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
input=""
for ((i=0; i<14; i++)); do
  input+="v$i "
done
read -a words <<< "$input"
[ "${#words[@]}" -eq 14 ] || fail "read -a argument count mismatch"
if (( 14 % 2 == 0 )); then
  [ "${words[0]}" = "v0" ] || fail "read array first"
else
  last=$((14 - 1))
  [ "${words[$last]}" = "v$last" ] || fail "read array last"
fi
echo PASS
exit 0
