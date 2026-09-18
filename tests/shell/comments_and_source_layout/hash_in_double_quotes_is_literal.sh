#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_in_double_quotes_is_literal
# A '#' character inside double quotes remains a literal hash character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg="hello # world"
[ "$msg" = "hello # world" ] || fail "double quoted hash: want 'hello # world', got [$msg]"
echo PASS
exit 0
