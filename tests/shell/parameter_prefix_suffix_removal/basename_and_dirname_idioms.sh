#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/basename_and_dirname_idioms
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
p=/usr/local/lib/file.tar.gz
[ "${p##*/}" = file.tar.gz ] || fail "basename: got [${p##*/}]"
[ "${p%/*}" = /usr/local/lib ] || fail "dirname: got [${p%/*}]"
b=${p##*/}
[ "${b%%.*}" = file ] || fail "stem: got [${b%%.*}]"
[ "${b#*.}" = tar.gz ] || fail "extension chain: got [${b#*.}]"
plain=file
[ "${plain%/*}" = file ] || fail "no slash leaves the value: got [${plain%/*}]"
echo PASS
exit 0
