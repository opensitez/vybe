#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/backslash_in_double_quotes_escapes_only_specials
# In "…", backslash is removed only before $ ` " \ and newline; elsewhere it stays.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="\$ \` \" \\ \a \n"
[ "$out" = '$ ` " \ \a \n' ] || fail "got [$out]"
echo PASS
exit 0
