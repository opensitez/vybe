#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_inside_subshell_group.sh
# A here-document can be defined and consumed cleanly inside a subshell compound command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sub_data=$(
    (
        cat <<EOF
subshell heredoc payload
EOF
    )
)
[ "$sub_data" = "subshell heredoc payload" ] || fail "subshell heredoc failed: got [$sub_data]"
echo PASS
exit 0
