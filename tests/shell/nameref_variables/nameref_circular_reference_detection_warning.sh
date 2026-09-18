#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_circular_reference_detection_warning
# Circular nameref definitions (a -> b -> a) emit a circular name reference warning to stderr.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
warn_output=$(
    (
        declare -n loop_a=loop_b
        declare -n loop_b=loop_a
        : "$loop_a"
    ) 2>&1
)
case "$warn_output" in
    *"circular name reference"*) : ;;
    *) fail "circular nameref did not emit expected warning: got [$warn_output]" ;;
esac
echo PASS
exit 0
