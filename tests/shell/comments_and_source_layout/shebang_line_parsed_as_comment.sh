#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/shebang_line_parsed_as_comment
# The first line shebang #!... is parsed by the shell interpreter as an ordinary comment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
[ "$x" -eq 1 ] || fail "shebang execution failed"
echo PASS
exit 0
