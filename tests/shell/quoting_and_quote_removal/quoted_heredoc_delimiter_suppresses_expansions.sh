#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quoted_heredoc_delimiter_suppresses_expansions
# Quoting any part of a here-document delimiter suppresses parameter and command expansions in the body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=42
body=$(cat <<"EOF"
value is $x and `echo backtick`
EOF
)
[ "$body" = 'value is $x and `echo backtick`' ] || fail "quoted heredoc body: got [$body]"
echo PASS
exit 0
