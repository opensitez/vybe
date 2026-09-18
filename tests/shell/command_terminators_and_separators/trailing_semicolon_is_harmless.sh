#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/trailing_semicolon_is_harmless
# One trailing ; is fine; two separators with nothing between them is an
# empty command and therefore a syntax error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo a;)
[ "$out" = a ] || fail "single trailing ; got [$out]"
eval 'echo a; ; echo b' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "; ; (empty command) must be a syntax error"
echo PASS
exit 0
