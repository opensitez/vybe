#!/usr/bin/env bash
# vybe-test: bash/expansion_order_and_interactions/tilde_prefix_from_expansion_is_not_a_login_name
# Tilde expansion runs before parameter expansion, so ~$u is checked as the
# literal login name "$u", which does not exist; the ~ stays.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
u=someone
out=$(echo ~$u)
[ "$out" = '~someone' ] || fail "got [$out]"
echo PASS
exit 0
