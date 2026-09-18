#!/usr/bin/env bash
# vybe-test: bash/builtin_read_arrays/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
input=""
for ((i=0; i<7; i++)); do
  input+="v$i "
done
read -a words <<< "$input"
[ "${#words[@]}" -eq 7 ] || fail "read -a argument count mismatch"
if (( 7 % 2 == 0 )); then
  [ "${words[0]}" = "v0" ] || fail "read array first"
else
  last=$((7 - 1))
  [ "${words[$last]}" = "v$last" ] || fail "read array last"
fi
echo PASS
exit 0
