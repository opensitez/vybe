#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_dash_f_displays_function_body
# The 'declare -f func_name' outputs the full function definition including its body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
my_display_fn() { printf 'payload\n'; }
body=$(declare -f my_display_fn)
case "$body" in
    *"my_display_fn ()"*|*"my_display_fn()"*) : ;;
    *) fail "declare -f failed to output function definition: got [$body]" ;;
esac
case "$body" in
    *"payload"*) : ;;
    *) fail "declare -f output missing function body: got [$body]" ;;
esac
echo PASS
exit 0
