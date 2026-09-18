#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/quoted_tabs_preserved_verbatim
# Tab characters enclosed in quotes are preserved as literal tab characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tab_str="col1	col2"
tab_len=${#tab_str}
[ "$tab_len" -eq 9 ] || fail "tab_str length: want 9, got $tab_len"
[ "$tab_str" = $'col1\tcol2' ] || fail "tab character mismatch: got [$tab_str]"
echo PASS
exit 0
