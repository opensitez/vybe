#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_lazy_evaluation_when_unset
# When the variable is unset, the alternate word is not evaluated (command substitutions are skipped).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_target
executed=0
res="${missing_target:+$(executed=1; echo 'should_not_run')}"
[ -z "$res" ] || fail "expansion should be empty: got [$res]"
[ "$executed" -eq 0 ] || fail "alternate word was eagerly evaluated"
echo PASS
exit 0
