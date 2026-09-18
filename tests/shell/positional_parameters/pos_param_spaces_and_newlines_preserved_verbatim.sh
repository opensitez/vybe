#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_spaces_and_newlines_preserved_verbatim
# Positional parameters containing spaces and embedded newlines are preserved verbatim.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline_arg="hello"$'\n'"world"
set -- "space arg" "$multiline_arg"
[ "$1" = "space arg" ] || fail "spaced argument corrupted: got [$1]"
[ "$2" = "$multiline_arg" ] || fail "multiline argument corrupted: got [$2]"
echo PASS
exit 0
