#!/usr/bin/env bash
# vybe-test: bash/builtin_read_arrays/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
input=""
for ((i=0; i<1; i++)); do
  input+="v$i "
done
read -a words <<< "$input"
[ "${#words[@]}" -eq 1 ] || fail "read -a argument count mismatch"
if (( 1 % 2 == 0 )); then
  [ "${words[0]}" = "v0" ] || fail "read array first"
else
  last=$((1 - 1))
  [ "${words[$last]}" = "v$last" ] || fail "read array last"
fi
echo PASS
exit 0
