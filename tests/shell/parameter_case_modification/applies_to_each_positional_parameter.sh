#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/applies_to_each_positional_parameter
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- ab cd EF
out="${@^}"
[ "$out" = "Ab Cd EF" ] || fail "^ on \$@: got [$out]"
out="${*,,}"
[ "$out" = "ab cd ef" ] || fail ",, on \$*: got [$out]"
echo PASS
exit 0
