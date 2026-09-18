#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/bang_negates_pipeline_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
! false; a=$?
! true; b=$?
! ! true; c=$?
! false | true; d=$?
[ "$a" = 0 ] || fail "! false want 0 got $a"
[ "$b" = 1 ] || fail "! true want 1 got $b"
[ "$c" = 0 ] || fail "! ! true want 0 got $c"
[ "$d" = 1 ] || fail "! applies to the whole pipeline, want 1 got $d"
echo PASS
exit 0
