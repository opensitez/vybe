#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_in_read_with_r_flag_preserved
# With -r, the read builtin disables backslash escaping and preserves backslashes literally.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r val <<< 'hello\ world'
[ "$val" = 'hello\ world' ] || fail "read -r: want 'hello\ world', got [$val]"
echo PASS
exit 0
