#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/expanded_word_is_never_an_assignment
# Assignment words are recognized before expansion, so a value of the form
# name=value produced by expansion is executed as a command name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset a
x='a=b'
$x 2>/dev/null; st=$?
[ "$st" -eq 127 ] || fail "want command-not-found 127 got $st"
[ -z "${a+set}" ] || fail "a must not have been assigned"
echo PASS
exit 0
