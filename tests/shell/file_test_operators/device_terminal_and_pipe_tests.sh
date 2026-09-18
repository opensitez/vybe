#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/device_terminal_and_pipe_tests
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -c /dev/null ] || fail "-c on /dev/null"
[ -f /dev/null ] && fail "/dev/null is not a regular file"
[ -e /dev/null ] && [ -w /dev/null ] || fail "-e/-w on /dev/null"
( [ -t 0 ] ) </dev/null && fail "-t 0 with stdin from /dev/null must be false"
: | { [ -p /dev/fd/0 ] || fail "-p on a pipe stdin"; }
[ -p /dev/null ] && fail "-p on a device"
echo PASS
exit 0
