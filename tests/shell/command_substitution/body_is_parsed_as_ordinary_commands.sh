#!/usr/bin/env bash
# vybe-test: bash/command_substitution/body_is_parsed_as_ordinary_commands
# A ) belonging to a case clause and a # comment inside $( ) do not end the
# substitution; the body is parsed like any other command list.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(case y in y) echo matched;; esac)
[ "$out" = matched ] || fail "case inside: got [$out]"
out=$( # a comment with ) inside
echo after-comment)
[ "$out" = after-comment ] || fail "comment inside: got [$out]"
echo PASS
exit 0
