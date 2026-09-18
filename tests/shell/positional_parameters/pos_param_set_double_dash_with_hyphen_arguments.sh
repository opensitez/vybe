#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_set_double_dash_with_hyphen_arguments
# The '--' delimiter prevents arguments starting with '-' or '+' from being treated as set options.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "-e" "-x" "+v" "--flag"
[ "$#" -eq 4 ] || fail "parameter count: want 4, got $#"
[ "$1" = "-e" ] || fail "param 1: want '-e', got [$1]"
[ "$2" = "-x" ] || fail "param 2: want '-x', got [$2]"
[ "$3" = "+v" ] || fail "param 3: want '+v', got [$3]"
[ "$4" = "--flag" ] || fail "param 4: want '--flag', got [$4]"
echo PASS
exit 0
