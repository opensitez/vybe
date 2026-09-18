#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_funcname_call_stack_variable
# The FUNCNAME array variable tracks the current function call stack frame.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inner_stack_check() {
    [ "${FUNCNAME[0]}" = "inner_stack_check" ] || fail "FUNCNAME[0]: want 'inner_stack_check', got [${FUNCNAME[0]}]"
    [ "${FUNCNAME[1]}" = "outer_stack_caller" ] || fail "FUNCNAME[1]: want 'outer_stack_caller', got [${FUNCNAME[1]}]"
}
outer_stack_caller() {
    inner_stack_check
}
outer_stack_caller
echo PASS
exit 0
