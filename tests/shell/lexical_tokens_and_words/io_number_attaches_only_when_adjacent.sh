#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/io_number_attaches_only_when_adjacent
# "echo 1>&2" makes 1 the descriptor of the redirection; "echo 1 >&2" passes
# 1 as an argument.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
adjacent=$( { echo 1>&2; } 2>&1 )
[ -z "$adjacent" ] || fail "adjacent digit must be an fd, got output [$adjacent]"
spaced=$( { echo 1 >&2; } 2>&1 )
[ "$spaced" = "1" ] || fail "spaced digit must be an argument, got [$spaced]"
echo PASS
exit 0
