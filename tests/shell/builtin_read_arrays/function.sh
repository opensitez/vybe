#!/usr/bin/env bash
# vybe-test: bash/builtin_read_arrays/function
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
input=""
for ((i=0; i<2; i++)); do
  input+="v$i "
done
read -a words <<< "$input"
[ "${#words[@]}" -eq 2 ] || fail "read -a argument count mismatch"
if (( 2 % 2 == 0 )); then
  [ "${words[0]}" = "v0" ] || fail "read array first"
else
  last=$((2 - 1))
  [ "${words[$last]}" = "v$last" ] || fail "read array last"
fi
echo PASS
exit 0
