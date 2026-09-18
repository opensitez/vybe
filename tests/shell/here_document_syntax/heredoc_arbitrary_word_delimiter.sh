#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_arbitrary_word_delimiter
# Any non-empty identifier token can serve as a here-document delimiter.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<__SPECIAL_DELIMITER_TOKEN__
content inside custom delimiter
__SPECIAL_DELIMITER_TOKEN__
)
[ "$doc" = "content inside custom delimiter" ] || fail "custom delimiter failed: got [$doc]"
echo PASS
exit 0
