#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/variable_identifiers_are_case_sensitive
# Variable identifiers differing only in letter case represent distinct storage locations.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target="lower"
Target="title"
TARGET="upper"
[ "$target" = "lower" ] || fail "target: want 'lower', got [$target]"
[ "$Target" = "title" ] || fail "Target: want 'title', got [$Target]"
[ "$TARGET" = "upper" ] || fail "TARGET: want 'upper', got [$TARGET]"
echo PASS
exit 0
