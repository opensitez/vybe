#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_process_substitution_input
# The <( ... ) construct creates an asynchronous process whose output acts as a readable file descriptor.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r line < <(printf 'streamed_line\n')
[ "$line" = "streamed_line" ] || fail "process substitution input: got [$line]"
echo PASS
exit 0
