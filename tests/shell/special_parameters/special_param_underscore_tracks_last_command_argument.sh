#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_underscore_tracks_last_command_argument
# The $_ parameter expands to the last argument of the previously executed command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: "first" "second" "third"
last_arg="$_"
[ "$last_arg" = "third" ] || fail "\$_ did not match last argument: want 'third', got [$last_arg]"
echo PASS
exit 0
