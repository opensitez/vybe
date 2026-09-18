#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_prefix_matching_with_custom_ifs
# The ${!prefix*} expansion joins matching variable names using the first character of IFS.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pfx_one=1
pfx_two=2
IFS=':'
joined="${!pfx_*}"
case "$joined" in
    *"pfx_one:pfx_two"*|*"pfx_two:pfx_one"*) : ;;
    *) fail "prefix match did not join with IFS=':': got [$joined]" ;;
esac
echo PASS
exit 0
