#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/heredoc_nested_within_command_substitution
# A here-document defined inside $(...) command substitution is parsed cleanly to its delimiter.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(cat <<EOF
nested heredoc content
EOF
)
[ "$out" = "nested heredoc content" ] || fail "nested heredoc in cmdsub: got [$out]"
echo PASS
exit 0
