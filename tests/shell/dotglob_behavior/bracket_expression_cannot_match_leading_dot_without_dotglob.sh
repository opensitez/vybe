#!/usr/bin/env bash
# vybe-test: bash/dotglob_behavior/bracket_expression_cannot_match_leading_dot_without_dotglob
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > .hidden
[ "$(echo [.]*)" = '[.]*' ] || fail "default: got [$(echo [.]*)]"
shopt -s dotglob
[ "$(echo [.]*)" = .hidden ] || fail "dotglob: got [$(echo [.]*)]"
echo PASS
exit 0
