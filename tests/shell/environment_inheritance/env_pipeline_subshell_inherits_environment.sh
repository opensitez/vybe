#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_pipeline_subshell_inherits_environment
# Pipeline stages running in subshells inherit the parent's environment variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export PIPELINE_VAR="stream_token"
piped_read=$(
    printf 'dummy\n' | {
        read -r _
        printf '%s\n' "$PIPELINE_VAR"
    }
)
[ "$piped_read" = "stream_token" ] || fail "pipeline element failed to inherit environment: got [$piped_read]"
echo PASS
exit 0
