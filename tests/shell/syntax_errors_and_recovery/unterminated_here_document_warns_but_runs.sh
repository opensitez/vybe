#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unterminated_here_document_warns_but_runs
# A here-document cut off by end-of-file is a warning, not an error: every
# remaining line is the body, the command runs, and the shell exits 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$("$BASH" -c $'while read -r l; do echo "line:$l"; done <<END\nhello\nworld' 2>&1); st=$?
[ "$st" -eq 0 ] || fail "want exit 0 got $st"
[[ $out == *"delimited by end-of-file"* ]] || fail "want warning got [$out]"
[[ $out == *"line:hello"*"line:world"* ]] || fail "both lines are body, got [$out]"
echo PASS
exit 0
