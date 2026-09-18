#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_case_sensitivity_in_environment_keys
# Child processes receive and distinguish distinct environment variables differing only by letter case.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export KeyName="mixed"
export KEYNAME="upper"
export keyname="lower"
res=$( "$BASH" -c 'printf "%s|%s|%s\n" "$KeyName" "$KEYNAME" "$keyname"' )
[ "$res" = "mixed|upper|lower" ] || fail "case-sensitivity corrupted in child process: got [$res]"
echo PASS
exit 0
