#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/function_definition_syntax_variants
# Bash recognizes three syntactic forms for function definitions: 'f()', 'function f', and 'function f()'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
f1() { printf 'form1\n'; }
function f2 { printf 'form2\n'; }
function f3() { printf 'form3\n'; }
[ "$(f1)" = "form1" ] || fail "f1: want 'form1', got [$(f1)]"
[ "$(f2)" = "form2" ] || fail "f2: want 'form2', got [$(f2)]"
[ "$(f3)" = "form3" ] || fail "f3: want 'form3', got [$(f3)]"
echo PASS
exit 0
