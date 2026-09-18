#!/usr/bin/env bash
# vybe-test: bash/process_substitution/input_form_provides_a_readable_path
# <(list) expands to a /dev/fd path from which list's output can be read.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
path=$(echo <(true))
[[ $path == /dev/fd/* ]] || fail "want a /dev/fd path got [$path]"
read -r line < <(echo hello)
[ "$line" = hello ] || fail "want [hello] got [$line]"
echo PASS
exit 0
