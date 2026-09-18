#!/usr/bin/env bash
# vybe-test: bash/expansion_order_and_interactions/redirection_target_is_expanded_but_not_split
# A redirection word must expand to exactly one word: a value with spaces is
# an "ambiguous redirect" unless it is quoted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
f='x y'
msg=$( { echo hi > $f; } 2>&1 ); st=$?
[ "$st" -ne 0 ] || fail "unquoted must fail"
[[ $msg == *"ambiguous redirect"* ]] || fail "got [$msg]"
[ ! -e x ] || fail "no file named x may be created"
echo hi > "$f"
[ "$(<"x y")" = hi ] || fail "quoted target must work"
echo PASS
exit 0
