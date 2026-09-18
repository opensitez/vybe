#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_appends_trailing_newline
# The here-string operator <<< supplies its word followed by a terminating newline character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IFS= read -r -d '' captured <<< "hello"
[ "${#captured}" -eq 6 ] || fail "captured length: want 6 (5 chars + newline), got ${#captured}"
[ "$captured" = $'hello\n' ] || fail "trailing newline missing in here-string"
echo PASS
exit 0
