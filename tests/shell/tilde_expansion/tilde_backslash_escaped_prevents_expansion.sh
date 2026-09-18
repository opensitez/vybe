#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_backslash_escaped_prevents_expansion
# Escaping a tilde with a backslash prevents tilde expansion, yielding a literal '~'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
HOME="/custom/home"
res=\~
[ "$res" = "~" ] || fail "backslash escaped tilde was expanded: got [$res]"
res_path=\~/path
[ "$res_path" = "~/path" ] || fail "backslash escaped \~/path was expanded: got [$res_path]"
echo PASS
exit 0
