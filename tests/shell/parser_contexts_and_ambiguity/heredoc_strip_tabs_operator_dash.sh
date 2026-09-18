#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/heredoc_strip_tabs_operator_dash
# The <<- redirection operator strips leading tab characters from lines and delimiter, but preserves spaces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<-EOF
	tab_stripped
	  tab_then_spaces_preserved
	EOF
)
expected="tab_stripped"$'\n'"  tab_then_spaces_preserved"
[ "$doc" = "$expected" ] || fail "heredoc <<- stripping: got [$doc]"
echo PASS
exit 0
