#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_double_quoted_prevents_expansion
# Double-quoting a tilde prevents tilde expansion, treating '~' as literal text.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
HOME="/custom/home"
res="~"
[ "$res" = "~" ] || fail "double quoted tilde was expanded: got [$res]"
res_path="~/dir"
[ "$res_path" = "~/dir" ] || fail "double quoted ~/dir was expanded: got [$res_path]"
echo PASS
exit 0
