#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_set_double_dash_initialization
# The 'set --' command populates positional parameters in left-to-right order.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "alpha" "beta" "gamma"
[ "$#" -eq 3 ] || fail "parameter count: want 3, got $#"
[ "$1" = "alpha" ] || fail "param 1: want 'alpha', got [$1]"
[ "$2" = "beta" ] || fail "param 2: want 'beta', got [$2]"
[ "$3" = "gamma" ] || fail "param 3: want 'gamma', got [$3]"
echo PASS
exit 0
