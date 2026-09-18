#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_trace_attribute_on_function_dash_t
# The 'declare -t' flag applies the trace attribute to a function, visible in declare -p/-F.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
traced_function() { :; }
declare -t -f traced_function
desc=$(declare -F)
case "$desc" in
    *"declare -ft traced_function"*) : ;;
    *) fail "trace attribute missing from declare -F output: got [$desc]" ;;
esac
echo PASS
exit 0
