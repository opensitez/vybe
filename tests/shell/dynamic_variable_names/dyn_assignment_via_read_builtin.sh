#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_assignment_via_read_builtin
# The read builtin accepts a dynamic variable identifier and populates it from standard input.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
dest_var="read_destination_entry"
read -r "$dest_var" <<< "incoming_record"
[ "$read_destination_entry" = "incoming_record" ] || fail "read dynamic assignment failed"
echo PASS
exit 0
