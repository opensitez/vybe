#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_strip_tabs_allows_indented_delimiter
# With <<-, the closing delimiter can be indented with leading tabs.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<-EOF
	body_line
	EOF
)
[ "$doc" = "body_line" ] || fail "indented delimiter with tabs failed: got [$doc]"
echo PASS
exit 0
