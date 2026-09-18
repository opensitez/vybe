#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_as_default_read_record_delimiter
# By default, the read builtin consumes characters up to the next newline as one input record.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    read -r rec1
    read -r rec2
} <<< $'record_one\nrecord_two'
[ "$rec1" = "record_one" ] || fail "rec1: want 'record_one', got [$rec1]"
[ "$rec2" = "record_two" ] || fail "rec2: want 'record_two', got [$rec2]"
echo PASS
exit 0
