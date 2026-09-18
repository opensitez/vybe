#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_process_substitution_output
# The >( ... ) construct creates an asynchronous process whose input acts as a writable file descriptor.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="initial"
tmp_out=""
# We can use cat through tee to write to variable or output
printf 'written_data\n' > >(read -r val; [ "$val" = "written_data" ] || exit 1)
wait
echo PASS
exit 0
