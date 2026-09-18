#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_in_single_quotes_is_literal
# A '#' character inside single quotes remains a literal hash character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg='single # quote'
[ "$msg" = "single # quote" ] || fail "single quoted hash: want 'single # quote', got [$msg]"
echo PASS
exit 0
