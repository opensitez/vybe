#!/usr/bin/env bash
# vybe-test: bash/builtin_read_arrays/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
input=""
for ((i=0; i<4; i++)); do
  input+="v$i "
done
read -a words <<< "$input"
[ "${#words[@]}" -eq 4 ] || fail "read -a argument count mismatch"
if (( 4 % 2 == 0 )); then
  [ "${words[0]}" = "v0" ] || fail "read array first"
else
  last=$((4 - 1))
  [ "${words[$last]}" = "v$last" ] || fail "read array last"
fi
echo PASS
exit 0
