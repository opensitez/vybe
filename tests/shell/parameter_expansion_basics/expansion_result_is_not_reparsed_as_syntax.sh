#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/expansion_result_is_not_reparsed_as_syntax
# Operators, redirections and quotes that come out of an expansion are plain
# characters in the word; they never become shell syntax.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
x='a;b'
[ "$(count $x)" = 1 ] || fail "; from expansion must not split commands"
[ "$(echo $x)" = 'a;b' ] || fail "got [$(echo $x)]"
r='>nonexistent_dir_zz/file'
out=$(echo $r)
[ "$out" = '>nonexistent_dir_zz/file' ] || fail "> from expansion must not redirect, got [$out]"
q="'a b'"
[ "$(count $q)" = 2 ] || fail "quotes from expansion do not group words"
echo PASS
exit 0
