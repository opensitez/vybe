#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_embeds_here_document
# A function body can contain here-documents that properly read parameters and format text.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
build_entry() {
    cat <<ENTRY
item=$1
status=$2
ENTRY
}
res=$(build_entry "pkg1" "installed")
expected="item=pkg1"$'\n'"status=installed"
[ "$res" = "$expected" ] || fail "function heredoc failed: got [$res]"
echo PASS
exit 0
