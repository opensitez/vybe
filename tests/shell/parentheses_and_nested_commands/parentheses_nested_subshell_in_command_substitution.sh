#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_nested_subshell_in_command_substitution
# An explicit subshell nested inside command substitution $( ( ... ) ) executes with extra subshell depth.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$( (printf 'from_inner_subshell\n') )
[ "$out" = "from_inner_subshell" ] || fail "nested subshell in cmdsub: got [$out]"
echo PASS
exit 0
