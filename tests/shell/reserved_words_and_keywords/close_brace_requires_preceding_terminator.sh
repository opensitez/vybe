#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/close_brace_requires_preceding_terminator
# { list } needs ; or newline before }, otherwise } is an argument.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$({ echo a; })
[ "$out" = a ] || fail "with semicolon: got [$out]"
out=$({ echo a
})
[ "$out" = a ] || fail "with newline: got [$out]"
eval '{ echo a }' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "missing terminator before } must be a syntax error"
echo PASS
exit 0
