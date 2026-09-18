#!/usr/bin/env bash
# vybe-test: bash/builtin_getopts_missing_arguments/repeat
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- -a "ok"
OPTIND=1
getopts ":a:b:" opt "$@"
[ "$?" -eq 0 ] || fail "valid option with argument failed"
set -- -b
OPTIND=1
getopts ":a:b:" opt "$@"
if [ "$?" -ne 0 ]; then
  :
else
  fail "missing option argument should fail"
fi
echo PASS
exit 0
