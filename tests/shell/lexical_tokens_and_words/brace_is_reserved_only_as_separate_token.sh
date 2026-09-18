#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/brace_is_reserved_only_as_separate_token
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo {a})
[ "$out" = "{a}" ] || fail "want [{a}] got [$out]"
out=$(echo })
[ "$out" = "}" ] || fail "want [}] got [$out]"
out=$({ echo grouped; })
[ "$out" = "grouped" ] || fail "want [grouped] got [$out]"
echo PASS
exit 0
