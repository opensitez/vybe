#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_argument_parsing_while_loop
# The idiomatic 'while [ $# -gt 0 ]; do ... shift; done' pattern processes all arguments sequentially.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "-a" "foo" "-b" "bar"
parsed_pairs=""
while [ "$#" -gt 0 ]; do
    opt="$1"
    val="$2"
    parsed_pairs+="$opt=$val;"
    shift 2
done
[ "$#" -eq 0 ] || fail "parameters remaining after while loop: got $#"
[ "$parsed_pairs" = "-a=foo;-b=bar;" ] || fail "parsed pairs mismatch: got [$parsed_pairs]"
echo PASS
exit 0
