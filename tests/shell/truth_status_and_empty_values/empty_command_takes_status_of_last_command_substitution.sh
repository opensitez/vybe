#!/usr/bin/env bash
# vybe-test: bash/truth_status_and_empty_values/empty_command_takes_status_of_last_command_substitution
# When every word expands to nothing, the command's status is that of the
# last command substitution performed, or 0 if there was none.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
$(false); st=$?
[ "$st" -eq 1 ] || fail "\$(false): want 1 got $st"
$(exit 3) $(exit 4); st=$?
[ "$st" -eq 4 ] || fail "two substitutions: want 4 got $st"
e=
$e; st=$?
[ "$st" -eq 0 ] || fail "plain empty expansion: want 0 got $st"
if $(false); then r=t; else r=f; fi
[ "$r" = f ] || fail "if \$(false) must take the else branch"
echo PASS
exit 0
