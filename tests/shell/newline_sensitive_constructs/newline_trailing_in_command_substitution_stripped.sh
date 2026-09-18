#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_trailing_in_command_substitution_stripped
# Command substitution $( ... ) strips all terminating trailing newline characters from output.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(printf 'result\n\n\n')
[ "$out" = "result" ] || fail "command substitution failed to strip trailing newlines: got [$out]"
[ "${#out}" -eq 6 ] || fail "length should be 6 without trailing newlines, got ${#out}"
echo PASS
exit 0
