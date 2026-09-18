#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_expands_command_substitution
# Command substitutions $( ... ) are evaluated and their output inserted into the here-string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r result <<< "output: $(printf 'generated_stream')"
[ "$result" = "output: generated_stream" ] || fail "command substitution in here-string: got [$result]"
echo PASS
exit 0
