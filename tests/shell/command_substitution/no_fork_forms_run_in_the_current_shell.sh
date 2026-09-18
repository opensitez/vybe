#!/usr/bin/env bash
# vybe-test: bash/command_substitution/no_fork_forms_run_in_the_current_shell
# bash 5.3: ${ list; } captures stdout like $( ) but runs list in the current
# shell, so its assignments persist; ${| list; } yields the value of REPLY.
# exit inside either form exits the shell itself.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
y=${ x=2; printf 'hi\n\n'; }
[ "$y" = hi ] || fail "captured output: want [hi] got [$y]"
[ "$x" = 2 ] || fail "assignment must persist, x=$x"
y=${ false; }; st=$?
[ "$st" -eq 1 ] || fail "status of the list: want 1 got $st"
z=${| REPLY=reply-value; other=set; }
[ "$z" = reply-value ] || fail "\${| } yields REPLY, got [$z]"
[ "$other" = set ] || fail "\${| } side effect must persist"
( y=${ exit 3; }; echo unreachable ); st=$?
[ "$st" -eq 3 ] || fail "exit inside must exit the shell with 3, got $st"
echo PASS
exit 0
