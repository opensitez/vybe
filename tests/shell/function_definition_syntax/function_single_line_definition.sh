#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_single_line_definition
# A function can be defined on a single line provided commands are terminated by semicolons.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inline_fn() { local msg="inline"; printf '%s\n' "$msg"; }
res=$(inline_fn)
[ "$res" = "inline" ] || fail "single line function failed: got [$res]"
echo PASS
exit 0
