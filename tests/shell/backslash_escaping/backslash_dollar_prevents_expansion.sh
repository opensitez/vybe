#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_dollar_prevents_expansion
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=value
out=$(echo \$x \${x} \$\(echo n\))
[ "$out" = '$x ${x} $(echo n)' ] || fail "got [$out]"
echo PASS
exit 0
