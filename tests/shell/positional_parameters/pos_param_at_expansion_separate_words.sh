#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_at_expansion_separate_words
# The "$@" parameter expansion expands each positional parameter to a separate quoted word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "arg with space" "second arg"
count_args() {
    [ "$#" -eq 2 ] || fail "expected 2 arguments, got $#"
    [ "$1" = "arg with space" ] || fail "arg 1 mismatch"
    [ "$2" = "second arg" ] || fail "arg 2 mismatch"
}
count_args "$@"
echo PASS
exit 0
