#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/local_default_from_optional_argument
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
greet() { local name=${1:-world}; echo "hello $name"; }
[ "$(greet)" = "hello world" ] || fail "default: got [$(greet)]"
[ "$(greet bob)" = "hello bob" ] || fail "given: got [$(greet bob)]"
[ "$(greet "")" = "hello world" ] || fail "empty argument uses default with :-"
echo PASS
exit 0
