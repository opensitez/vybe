#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_quoted_preserves_single_word
# Double-quoting the alternate value expansion preserves it as a single intact argument.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
flag="enabled"
count_words() {
    [ "$#" -eq 1 ] || fail "quoted alternate word preservation failed: want 1 word, got $#"
    [ "$1" = "one two three" ] || fail "content mismatch: got [$1]"
}
count_words "${flag:+one two three}"
echo PASS
exit 0
