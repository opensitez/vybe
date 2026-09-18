#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/transformation_of_unset_variable_is_empty
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset nothing
[ -z "${nothing@Q}" ] || fail "@Q: got [${nothing@Q}]"
[ -z "${nothing@a}" ] || fail "@a: got [${nothing@a}]"
[ -z "${nothing@U}" ] || fail "@U: got [${nothing@U}]"
echo PASS
exit 0
