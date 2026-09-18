#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/multiple_prefix_assignment_tokens
# Multiple prefix assignment tokens before a command assign variables in command environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
A=10 B=20 bash -c 'exit $((A + B == 30 ? 0 : 1))'
st=$?
[ "$st" -eq 0 ] || fail "multiple prefix assignments: want status 0, got $st"
echo PASS
exit 0
