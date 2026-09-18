#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/nested_compound_commands_execution
# Compound commands nest arbitrarily (e.g. for inside if inside a brace group).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
collected=""
{
    if true; then
        for ch in a b c; do
            case "$ch" in
                b) collected+="[B]" ;;
                *) collected+="$ch" ;;
            esac
        done
    fi
}
[ "$collected" = "a[B]c" ] || fail "nested compound commands: want 'a[B]c', got [$collected]"
echo PASS
exit 0
