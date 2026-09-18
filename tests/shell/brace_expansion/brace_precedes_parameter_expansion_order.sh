#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_precedes_parameter_expansion_order
# In Bash, brace expansion occurs before parameter expansion; range endpoints cannot be variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
max=5
set -- {1..$max}
# Because brace expansion occurs before variable expansion, {1..$max} does not form a valid range
[ "$#" -eq 1 ] || fail "count: want 1, got $#"
[ "$1" = "{1..5}" ] || fail "brace expansion ordering failed: got [$1]"
echo PASS
exit 0
