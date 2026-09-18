#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_array_expansion_star
# Expanding "${arr[*]}" in a here-string joins array elements with the first character of IFS.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( a b c )
IFS=':'
read -r joined <<< "${arr[*]}"
[ "$joined" = "a:b:c" ] || fail "array star expansion with IFS: want 'a:b:c', got [$joined]"
echo PASS
exit 0
