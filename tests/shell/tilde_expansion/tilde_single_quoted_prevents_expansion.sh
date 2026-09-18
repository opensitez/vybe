#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_single_quoted_prevents_expansion
# Single-quoting a tilde prevents tilde expansion, treating '~' as literal text.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
HOME="/custom/home"
res='~'
[ "$res" = "~" ] || fail "single quoted tilde was expanded: got [$res]"
res_path='~/path'
[ "$res_path" = '~/path' ] || fail "single quoted ~/path was expanded: got [$res_path]"
echo PASS
exit 0
