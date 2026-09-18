#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_after_case_pattern_pipe
# Comments can appear between case clauses and following case pattern terminators ';;'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val="b"
match=""
case "$val" in
    a) # match first pattern
        match="found_a"
        ;; # terminate clause
    b) # match second pattern
        match="found_b"
        ;; # terminate clause
    *) # fallback default
        match="default"
        ;;
esac
[ "$match" = "found_b" ] || fail "case comments match: want 'found_b', got [$match]"
echo PASS
exit 0
