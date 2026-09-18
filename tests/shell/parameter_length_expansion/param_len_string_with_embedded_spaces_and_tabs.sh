#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_string_with_embedded_spaces_and_tabs
# The ${#var} expansion counts horizontal spaces and tab characters verbatim.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
space_tab_str="a "$'\t'" b"
[ "${#space_tab_str}" -eq 5 ] || fail "spaces and tab length: want 5, got ${#space_tab_str}"
echo PASS
exit 0
