#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_empty_body_yields_empty_string
# A here-document with no content lines between the opening line and delimiter produces an empty stream.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<EOF
EOF
)
[ -z "$doc" ] || fail "empty heredoc should produce empty string, got [$doc]"
echo PASS
exit 0
