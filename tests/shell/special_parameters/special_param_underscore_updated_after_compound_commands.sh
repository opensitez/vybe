#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_underscore_updated_after_compound_commands
# The $_ parameter is updated following the completion of compound commands and loops.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
for item in apple banana cherry; do
    : "$item"
done
[ "$_" = "cherry" ] || fail "\$_ after for loop: want 'cherry', got [$_]"
echo PASS
exit 0
