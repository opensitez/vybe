#!/usr/bin/env bash
# vybe-test: bash/variable_existence_tests/dash_v_on_positional_parameters
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- one ""
[[ -v 1 ]] || fail "\$1 is set"
[[ -v 2 ]] || fail "\$2 is set (to empty)"
[[ -v 3 ]] && fail "\$3 is not set"
echo PASS
exit 0
