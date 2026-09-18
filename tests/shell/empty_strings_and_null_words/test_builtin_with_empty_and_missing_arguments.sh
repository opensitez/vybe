#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/test_builtin_with_empty_and_missing_arguments
# [ ] is false, [ "" ] is false, [ -n ] is TRUE (one non-empty argument "-n").
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ ] && fail "[ ] must be false"
[ "" ] && fail "[ \"\" ] must be false"
[ -n ] || fail "[ -n ] with one argument is a non-empty string test on -n itself"
[ -z "" ] || fail "[ -z \"\" ] must be true"
[ -n "" ] && fail "[ -n \"\" ] must be false"
echo PASS
exit 0
