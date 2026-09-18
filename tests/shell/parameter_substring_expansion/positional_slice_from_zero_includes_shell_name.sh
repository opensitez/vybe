#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/positional_slice_from_zero_includes_shell_name
# ${@:0} starts at $0; ${@:1} is the same as "$@".
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
set -- a b c
n=$(count "${@:0}")
[ "$n" = 4 ] || fail "want 4 words got $n"
first() { echo "$1"; }
[ "$(first "${@:0:1}")" = "$0" ] || fail "first of \${@:0} must be \$0"
[ "$(count "${@:1}")" = 3 ] || fail "\${@:1} must be all positionals"
[ "${*:2:2}" = "b c" ] || fail "want [b c] got [${*:2:2}]"
echo PASS
exit 0
