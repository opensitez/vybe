#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_clean_invocation_with_explicit_variables
# 'env -i VAR=val command' provides a sanitized environment containing only explicitly passed variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export PARENT_SECRET="omit_me"
res=$( env -i ONLY_THIS="exact_payload" "$BASH" -c 'printf "%s|%s\n" "$ONLY_THIS" "$PARENT_SECRET"' )
[ "$res" = "exact_payload|" ] || fail "sanitized env failed: got [$res]"
echo PASS
exit 0
