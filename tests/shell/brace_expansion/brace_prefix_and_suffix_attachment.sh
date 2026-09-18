#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_prefix_and_suffix_attachment
# Preamble and postscript strings are prepended and appended to each expanded word in braces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- pre_{one,two,three}_post
[ "$#" -eq 3 ] || fail "count: want 3, got $#"
[ "$1" = "pre_one_post" ] || fail "word 1 mismatch: got [$1]"
[ "$2" = "pre_two_post" ] || fail "word 2 mismatch: got [$2]"
[ "$3" = "pre_three_post" ] || fail "word 3 mismatch: got [$3]"
echo PASS
exit 0
