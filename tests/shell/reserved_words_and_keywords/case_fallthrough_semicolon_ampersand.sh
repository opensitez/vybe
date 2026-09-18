#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/case_fallthrough_semicolon_ampersand
# The ;& operator in a case statement executes the command list of the subsequent clause without pattern matching.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=""
case "a" in
    a)
        acc+="1"
        ;&
    b)
        acc+="2"
        ;;
    c)
        acc+="3"
        ;;
esac
[ "$acc" = "12" ] || fail "case fallthrough with ;&: want '12', got [$acc]"
echo PASS
exit 0
