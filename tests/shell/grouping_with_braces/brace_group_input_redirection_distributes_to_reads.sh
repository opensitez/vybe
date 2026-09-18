#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_input_redirection_distributes_to_reads
# Supplying input redirection to a brace group streams records sequentially across internal read calls.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    read -r line_a
    read -r line_b
} <<< $'first_record\nsecond_record'
[ "$line_a" = "first_record" ] || fail "line_a: want 'first_record', got [$line_a]"
[ "$line_b" = "second_record" ] || fail "line_b: want 'second_record', got [$line_b]"
echo PASS
exit 0
