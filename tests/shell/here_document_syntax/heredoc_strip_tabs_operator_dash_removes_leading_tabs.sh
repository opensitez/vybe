#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_strip_tabs_operator_dash_removes_leading_tabs
# The <<- operator strips all leading tab characters from body lines.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<-EOF
	tab1
		tab2
EOF
)
expected="tab1"$'\n'"tab2"
[ "$doc" = "$expected" ] || fail "leading tabs not stripped: got [$doc]"
echo PASS
exit 0
