#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_disabled_via_set_plus_B
# Disabling brace expansion using 'set +B' prevents brace expansion, preserving braces literally.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set +B
set -- {1..3}
[ "$#" -eq 1 ] || fail "set +B count: want 1, got $#"
[ "$1" = "{1..3}" ] || fail "set +B failed to disable brace expansion: got [$1]"

set -- {a,b}
[ "$#" -eq 1 ] || fail "comma count under set +B: want 1, got $#"
[ "$1" = "{a,b}" ] || fail "set +B failed on comma expansion: got [$1]"
echo PASS
exit 0
