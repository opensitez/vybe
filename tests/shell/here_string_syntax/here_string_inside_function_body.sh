#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_inside_function_body
# A function can parse its argument using an internal here-string fed into read.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
parse_pair() {
    local key val
    read -r key val <<< "$1"
    printf '%s=%s\n' "$key" "$val"
}
out=$(parse_pair "user admin")
[ "$out" = "user=admin" ] || fail "function internal here-string failed: got [$out]"
echo PASS
exit 0
