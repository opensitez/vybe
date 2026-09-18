#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_pwd_updates_automatically_on_cd
# Changing directory updates the $PWD environment variable to the absolute canonical path.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
orig="$PWD"
cd /tmp
[ "$PWD" = "/tmp" ] || [ "$PWD" = "/private/tmp" ] || fail "PWD did not update: got [$PWD]"
cd "$orig"
[ "$PWD" = "$orig" ] || fail "failed to restore PWD: got [$PWD]"
echo PASS
exit 0
