#!/usr/bin/env bash
# vybe-test: bash/process_substitution/available_in_posix_mode
# Process substitution is a bash extension that stays enabled under set -o posix.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -o posix
read -r l < <(echo posix)
[ "$l" = posix ] || fail "got [$l]"
echo PASS
exit 0
