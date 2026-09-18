#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_malformed_expression_syntax_is_reported
# Invalid arithmetic tokenization in $(( )) is a parse error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'echo $((1 + * 2))' 2>&1); st=$?
[ "$st" -eq 1 ] || fail "want status 1, got $st"
[[ $msg == *"arithmetic syntax error"* || $msg == *"syntax error"* ]] || fail "unexpected message: [$msg]"
echo PASS
exit 0
