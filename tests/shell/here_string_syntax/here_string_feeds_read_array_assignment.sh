#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_feeds_read_array_assignment
# Feeding a here-string to 'read -a array' populates an indexed array with the split elements.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -ra items <<< "elem0 elem1 elem2 elem3"
[ "${#items[@]}" -eq 4 ] || fail "array length: want 4, got ${#items[@]}"
[ "${items[0]}" = "elem0" ] || fail "items[0]: want 'elem0', got [${items[0]}]"
[ "${items[3]}" = "elem3" ] || fail "items[3]: want 'elem3', got [${items[3]}]"
echo PASS
exit 0
