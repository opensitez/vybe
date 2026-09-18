#!/usr/bin/env bash
# vybe-test: bash/process_substitution/several_substitutions_as_separate_arguments
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
join() { local a b; read -r a < "$1"; read -r b < "$2"; echo "$a$b"; }
out=$(join <(echo x) <(echo y))
[ "$out" = xy ] || fail "want [xy] got [$out]"
echo PASS
exit 0
