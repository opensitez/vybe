#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_reenabled_via_set_minus_B
# Re-enabling brace expansion using 'set -B' restores normal brace expansion behavior.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set +B
set -- {a,b}
[ "$1" = "{a,b}" ] || fail "disabling failed"

set -B
set -- {a,b}
[ "$#" -eq 2 ] || fail "reenabling count: want 2, got $#"
[ "$1" = "a" ] && [ "$2" = "b" ] || fail "reenabling failed: got [$1], [$2]"
echo PASS
exit 0
