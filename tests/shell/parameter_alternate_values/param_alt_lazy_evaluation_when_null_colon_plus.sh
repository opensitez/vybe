#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_lazy_evaluation_when_null_colon_plus
# When the variable is null (empty), the alternate word in ${var:+word} is not evaluated.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_target=""
executed=0
res="${empty_target:+$(executed=1; echo 'should_not_run')}"
[ -z "$res" ] || fail "expansion should be empty: got [$res]"
[ "$executed" -eq 0 ] || fail "alternate word was eagerly evaluated for null variable"
echo PASS
exit 0
