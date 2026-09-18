#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/adjacent_empty_quotes_form_one_null_word
# Any quoted part, even empty, forces a word to exist; an unquoted empty
# expansion next to it adds nothing.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
e=
[ "$(count ""'')" = 1 ] || fail "\"\"'' want 1 got $(count ""'')"
[ "$(count ""$e)" = 1 ] || fail "\"\"\$e want 1 got $(count ""$e)"
[ "$(count $e$e)" = 0 ] || fail "\$e\$e want 0 got $(count $e$e)"
echo PASS
exit 0
