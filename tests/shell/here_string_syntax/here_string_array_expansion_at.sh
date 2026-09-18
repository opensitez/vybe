#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_array_expansion_at
# Expanding "${arr[@]}" in a here-string separates elements with spaces into a single stream.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "part 1" "part 2" )
read -r joined <<< "${arr[@]}"
[ "$joined" = "part 1 part 2" ] || fail "array at expansion in here-string failed: got [$joined]"
echo PASS
exit 0
