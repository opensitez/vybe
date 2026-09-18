#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/printf_treats_missing_and_empty_arguments_alike
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ "$(printf '[%s]' "")" = '[]' ] || fail "empty arg: got [$(printf '[%s]' "")]"
[ "$(printf '[%s]')" = '[]' ] || fail "missing arg: got [$(printf '[%s]')]"
[ "$(printf '[%d]')" = '[0]' ] || fail "missing numeric arg is 0: got [$(printf '[%d]')]"
echo PASS
exit 0
