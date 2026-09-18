#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/type_t_distinguishes_keyword_builtin_and_function
# 'type -t' resolves and categorizes an identifier into keyword, builtin, or function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
custom_func() { :; }
t_kw=$(type -t for)
t_bi=$(type -t cd)
t_fn=$(type -t custom_func)
[ "$t_kw" = "keyword" ] || fail "for category: want 'keyword', got [$t_kw]"
[ "$t_bi" = "builtin" ] || fail "cd category: want 'builtin', got [$t_bi]"
[ "$t_fn" = "function" ] || fail "custom_func category: want 'function', got [$t_fn]"
echo PASS
exit 0
