#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_escaped_quotes_inside_string
# Escaping double quotes inside a double-quoted assignment includes literal quote characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
escaped="She said, \"Hello, world!\""
[ "$escaped" = 'She said, "Hello, world!"' ] || fail "escaped double quotes failed: got [$escaped]"
echo PASS
exit 0
