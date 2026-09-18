#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_dash_p_lists_declarations
# The 'readonly -p' builtin outputs declarations starting with 'declare -r' or 'readonly'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly MY_UNIQUE_RO_FLAG="unique_flag_val"
output=$(readonly -p)
case "$output" in
    *MY_UNIQUE_RO_FLAG*) : ;;
    *) fail "readonly -p output did not contain declared readonly variable" ;;
esac
echo PASS
exit 0
