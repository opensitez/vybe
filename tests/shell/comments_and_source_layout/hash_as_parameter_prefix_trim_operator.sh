#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_as_parameter_prefix_trim_operator
# The '#' and '##' in ${var#pattern} and ${var##pattern} are prefix trim operators, not comments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
path="/usr/local/bin"
trim1=${path#*/}
trim2=${path##*/}
[ "$trim1" = "usr/local/bin" ] || fail "trim1: want 'usr/local/bin', got [$trim1]"
[ "$trim2" = "bin" ] || fail "trim2: want 'bin', got [$trim2]"
echo PASS
exit 0
