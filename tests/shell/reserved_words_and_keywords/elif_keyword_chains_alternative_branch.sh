#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/elif_keyword_chains_alternative_branch
# The elif reserved word chains an alternative condition in an if statement compound command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=2
branch=""
if [ "$x" -eq 1 ]; then
    branch="first"
elif [ "$x" -eq 2 ]; then
    branch="second"
else
    branch="third"
fi
[ "$branch" = "second" ] || fail "elif branch selection: want 'second', got [$branch]"
echo PASS
exit 0
