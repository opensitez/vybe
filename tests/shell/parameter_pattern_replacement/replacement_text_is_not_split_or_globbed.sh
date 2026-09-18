#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/replacement_text_is_not_split_or_globbed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
x=abc
r='x  y'
[ "$(count "${x/b/$r}")" = 1 ] || fail "quoted result is one word"
[ "${x/b/$r}" = 'ax  yc' ] || fail "spaces kept: got [${x/b/$r}]"
[ "${x/b/*}" = 'a*c' ] || fail "* in replacement is literal: got [${x/b/*}]"
[ "${x/b//}" = 'a/c' ] || fail "slash as replacement: got [${x/b//}]"
echo PASS
exit 0
