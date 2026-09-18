#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/case_statement_word_and_pattern_in_keyword
# The word 'in' can be used as the target word and pattern within a case statement without syntax collision.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=""
case in in
    in)
        res="matched_in"
        ;;
    *)
        res="fallback"
        ;;
esac
[ "$res" = "matched_in" ] || fail "case in in in: want 'matched_in', got [$res]"
echo PASS
exit 0
