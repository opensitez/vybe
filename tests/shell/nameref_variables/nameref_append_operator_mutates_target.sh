#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_append_operator_mutates_target
# Using the '+=' append operator on a nameref appends directly to the target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
buf="initial"
declare -n ref=buf
ref+="_appended"
[ "$buf" = "initial_appended" ] || fail "append through nameref failed: got [$buf]"
[ "$ref" = "initial_appended" ] || fail "reading appended nameref failed"
echo PASS
exit 0
