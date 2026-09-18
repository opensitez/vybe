#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_string_append_assignment_operator
# The '+=' operator appends a string to an existing scalar variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
buf="hello"
buf+=" world"
buf+="!"
[ "$buf" = "hello world!" ] || fail "append assignment failed: got [$buf]"
echo PASS
exit 0
