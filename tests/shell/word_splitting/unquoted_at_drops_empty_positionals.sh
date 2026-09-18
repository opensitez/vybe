#!/usr/bin/env bash
# vybe-test: bash/word_splitting/unquoted_at_drops_empty_positionals
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
set -- a "" b
[ "$(count $@)" = 2 ] || fail "unquoted \$@: want 2 got $(count $@)"
[ "$(count "$@")" = 3 ] || fail "quoted \"\$@\": want 3"
echo PASS
exit 0
