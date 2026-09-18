#!/usr/bin/env bash
# vybe-test: bash/expansion_order_and_interactions/command_substitution_output_is_globbed_when_unquoted
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > a.txt; : > b.txt
out=$(echo $(echo '*.txt'))
[ "$out" = "a.txt b.txt" ] || fail "unquoted: got [$out]"
out=$(echo "$(echo '*.txt')")
[ "$out" = '*.txt' ] || fail "quoted: got [$out]"
echo PASS
exit 0
