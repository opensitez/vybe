#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/backslash_escaped_newline_line_continuation
# A backslash immediately preceding a newline removes both characters, merging lines seamlessly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
composite_cmd="part1"\
"part2"
[ "$composite_cmd" = "part1part2" ] || fail "line continuation: want 'part1part2', got [$composite_cmd]"
echo PASS
exit 0
