#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_in_read_without_r_flag
# Without -r, the read builtin treats backslashes as escape characters and removes them.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read val <<< 'hello\ world'
[ "$val" = "hello world" ] || fail "read without -r: want 'hello world', got [$val]"
echo PASS
exit 0
