#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_overridden_by_subsequent_redirection
# When multiple standard input redirections appear on a command, the rightmost redirection wins.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r line <<< "first_input" <<< "second_input"
[ "$line" = "second_input" ] || fail "rightmost redirection should take precedence: got [$line]"
echo PASS
exit 0
