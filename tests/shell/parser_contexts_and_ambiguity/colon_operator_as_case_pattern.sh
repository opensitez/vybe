#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/colon_operator_as_case_pattern
# The colon ':' character acts as a valid literal pattern inside case statement clauses.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val=":"
match=""
case "$val" in
    :)
        match="colon_matched"
        ;;
    *)
        match="fallback"
        ;;
esac
[ "$match" = "colon_matched" ] || fail "case colon pattern: want 'colon_matched', got [$match]"
echo PASS
exit 0
