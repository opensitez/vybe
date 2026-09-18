#!/usr/bin/env bash
# vybe-test: bash/command_substitution/backticks_nest_only_with_escaping
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=`echo hi`
[ "$out" = hi ] || fail "simple backticks: got [$out]"
out="`echo quoted`"
[ "$out" = quoted ] || fail "inside double quotes: got [$out]"
out=`echo \`echo inner\``
[ "$out" = inner ] || fail "escaped nesting: got [$out]"
echo PASS
exit 0
