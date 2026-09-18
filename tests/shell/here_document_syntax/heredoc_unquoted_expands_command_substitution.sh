#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_unquoted_expands_command_substitution
# In an unquoted here-document, $(...) command substitutions are executed and substituted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<EOF
computed: $(printf 'generated_value')
EOF
)
[ "$doc" = "computed: generated_value" ] || fail "cmdsub in heredoc failed: got [$doc]"
echo PASS
exit 0
