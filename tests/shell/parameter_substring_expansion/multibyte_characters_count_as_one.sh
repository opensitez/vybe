#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/multibyte_characters_count_as_one
# Under a UTF-8 locale offsets and lengths are in characters, not bytes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export LC_ALL=en_US.UTF-8
x='héllo'
[ "${x:1:1}" = 'é' ] || fail "want [é] got [${x:1:1}]"
[ "${x:2}" = 'llo' ] || fail "want [llo] got [${x:2}]"
[ "${x: -1}" = 'o' ] || fail "want [o] got [${x: -1}]"
echo PASS
exit 0
