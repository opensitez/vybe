#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_negative_offset_from_end
# The syntax ${@: -1} with leading whitespace indexes positional parameters from the end.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "first" "middle" "last"
last_arg="${@: -1}"
[ "$last_arg" = "last" ] || fail "negative offset failed: want 'last', got [$last_arg]"
second_to_last="${@: -2:1}"
[ "$second_to_last" = "middle" ] || fail "negative offset with length failed: want 'middle', got [$second_to_last]"
echo PASS
exit 0
