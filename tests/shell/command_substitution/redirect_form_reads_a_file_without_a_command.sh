#!/usr/bin/env bash
# vybe-test: bash/command_substitution/redirect_form_reads_a_file_without_a_command
# $(< file) is a built-in shortcut for reading a file; a missing file gives
# an empty result and status 1 (the diagnostic is printed during expansion,
# so it must be silenced on an enclosing group, not on the assignment).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
exec 3<<<$'line one\nline two'
x=$(</dev/fd/3)
[ "$x" = $'line one\nline two' ] || fail "got [$x]"
{ y=$(<./no_such_file_zz); st=$?; } 2>/dev/null
[ -z "$y" ] || fail "missing file must give empty, got [$y]"
[ "$st" -eq 1 ] || fail "missing file status want 1 got $st"
echo PASS
exit 0
