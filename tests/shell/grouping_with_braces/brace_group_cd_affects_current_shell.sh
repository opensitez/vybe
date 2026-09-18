#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_cd_affects_current_shell
# Changing directory inside a brace group affects the working directory of the enclosing shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
orig="$PWD"
{
    cd /tmp
}
[ "$PWD" = "/tmp" ] || [ "$PWD" = "/private/tmp" ] || fail "cd inside brace group failed: got [$PWD]"
cd "$orig"
[ "$PWD" = "$orig" ] || fail "failed to restore directory: got [$PWD]"
echo PASS
exit 0
