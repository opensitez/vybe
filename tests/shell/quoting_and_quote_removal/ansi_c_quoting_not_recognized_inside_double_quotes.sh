#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/ansi_c_quoting_not_recognized_inside_double_quotes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="$'a\tb'"
[ "$out" = "\$'a\\tb'" ] || fail "want literal \$'a\\tb' got [$out]"
[ "${#out}" -eq 7 ] || fail "length want 7 got ${#out}"
echo PASS
exit 0
