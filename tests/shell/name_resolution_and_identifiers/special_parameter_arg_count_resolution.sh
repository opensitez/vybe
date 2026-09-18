#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/special_parameter_arg_count_resolution
# The special parameter identifier '#' resolves to the count of positional parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count_args() {
    [ "$#" -eq 4 ] || fail "arg count: want 4, got $#"
}
count_args one two three four
echo PASS
exit 0
