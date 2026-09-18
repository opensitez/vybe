#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_suppression_via_echo_dash_n
# The echo -n option suppresses the default trailing newline on output.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out1=$(echo "hello")
out2=$(echo -n "hello")
# Both get trailing newlines stripped by $(), but we can test by piping into wc -c
IFS= read -r -d '' len_without_nl < <(echo -n "hello")
IFS= read -r -d '' len_with_nl < <(echo "hello")
[ "${#len_with_nl}" -eq 6 ] || fail "len with nl: want 6, got ${#len_with_nl}"
[ "${#len_without_nl}" -eq 5 ] || fail "len without nl: want 5, got ${#len_without_nl}"
echo PASS
exit 0
