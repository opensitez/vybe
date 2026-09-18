#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_masks_command_substitution_exit_code
# In 'local var=$(cmd)', the exit status is 0 (from the local builtin itself), masking cmd's exit failure.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
mask_fn() {
    local res=$(exit 44)
    local st=$?
    # Because 'local' executed successfully, $? from 'local res=$(exit 44)' is 0
    [ "$st" -eq 0 ] || fail "local should mask subshell exit code: want 0, got $st"
}
mask_fn
echo PASS
exit 0
