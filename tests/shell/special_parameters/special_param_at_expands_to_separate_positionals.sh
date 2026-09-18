#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_at_expands_to_separate_positionals
# The "$@" parameter expands each positional parameter to an individual word preserving spaces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "part 1" "part 2"
verify_args() {
    [ "$#" -eq 2 ] || fail "arg count: want 2, got $#"
    [ "$1" = "part 1" ] && [ "$2" = "part 2" ] || fail "arg content mismatch"
}
verify_args "$@"
echo PASS
exit 0
