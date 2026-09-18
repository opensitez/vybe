#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_empty_elements_in_comma_list
# Empty elements within brace expansion attach to preamble/postscript and quoted empty strings preserve words.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
# With preamble, empty element yields preamble alone
set -- pre{a,,c}
[ "$#" -eq 3 ] || fail "preamble empty element count: want 3, got $#"
[ "$1" = "prea" ] && [ "$2" = "pre" ] && [ "$3" = "prec" ] || fail "empty middle element failed: got [$1], [$2], [$3]"

# With explicit empty quotes, an empty argument is preserved
set -- {a,"",c}
[ "$#" -eq 3 ] || fail "quoted empty element count: want 3, got $#"
[ "$1" = "a" ] && [ -z "$2" ] && [ "$3" = "c" ] || fail "quoted empty element failed: got [$1], [$2], [$3]"
echo PASS
exit 0
