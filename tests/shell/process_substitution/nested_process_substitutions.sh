#!/usr/bin/env bash
# vybe-test: bash/process_substitution/nested_process_substitutions
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r l < <(read -r m < <(echo deep); echo "in:$m")
[ "$l" = in:deep ] || fail "got [$l]"
echo PASS
exit 0
