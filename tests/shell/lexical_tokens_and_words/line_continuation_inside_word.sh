#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/line_continuation_inside_word
# Backslash-newline is removed during tokenizing, so it can split a single
# word across lines.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=hel\
lo
[ "$x" = "hello" ] || fail "want [hello] got [$x]"
[ "${#x}" -eq 5 ] || fail "length: want 5 got ${#x}"
echo PASS
exit 0
