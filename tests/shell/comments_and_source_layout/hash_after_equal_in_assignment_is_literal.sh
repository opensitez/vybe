#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_after_equal_in_assignment_is_literal
# An unquoted '#' immediately following '=' in variable assignment is literal because it is not at the start of a word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tag=#status
[ "$tag" = "#status" ] || fail "hash after equal: want '#status', got [$tag]"
echo PASS
exit 0
