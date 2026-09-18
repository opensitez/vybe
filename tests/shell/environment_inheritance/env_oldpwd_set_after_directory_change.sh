#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_oldpwd_set_after_directory_change
# After changing directory via cd, $OLDPWD is populated with the previous working directory.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
start_dir="$PWD"
cd /tmp
[ "$OLDPWD" = "$start_dir" ] || fail "OLDPWD was not set to initial directory: got [$OLDPWD]"
cd "$start_dir"
echo PASS
exit 0
