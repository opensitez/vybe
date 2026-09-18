#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/keyword_as_argument_is_literal
# Reserved words are recognized only as the first word of a command (or after
# certain other reserved words).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo for do done while until case esac function select)
[ "$out" = "for do done while until case esac function select" ] || fail "got [$out]"
echo PASS
exit 0
