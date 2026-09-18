#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/empty_command_substitution_body_is_error
# $() with nothing inside is fine (empty string); a lone ( ) is not.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$()
[ -z "$x" ] || fail "empty \$() should expand to nothing"
eval '( )' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "empty subshell must be a syntax error"
echo PASS
exit 0
