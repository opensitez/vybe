#!/usr/bin/env bash
# vybe-test: bash/bash_function_metadata/metadata_probe
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=7
function_meta(){ local value=$IDX; return 0; }
meta_name=$(declare -F function_meta)
[[ $meta_name == *"function_meta"* ]] || fail "declare -F missing function name"
meta_body=$(declare -f function_meta)
[[ $meta_body == *"local value="* ]] || fail "declare -f should include function body"
unset -f function_meta
[[ -z $(type -t function_meta 2>/dev/null) ]] || fail "function metadata function was not removed"
echo PASS
exit 0
