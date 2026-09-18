#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_inside_function_body
# A function body can contain a here-document that interpolates function positional parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
render_msg() {
    cat <<EOF
Recipient: $1
Status: $2
EOF
}
res=$(render_msg "Bob" "Active")
expected="Recipient: Bob"$'\n'"Status: Active"
[ "$res" = "$expected" ] || fail "function heredoc failed: got [$res]"
echo PASS
exit 0
