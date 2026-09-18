#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_E_expands_escape_sequences
# The ${var@E} parameter transformation expands backslash escape sequences as with $'...' quoting.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
raw_escapes='line1\nline2\ttab'
expanded="${raw_escapes@E}"
case "$expanded" in
    *"
"*) : ;;
    *) fail "\${var@E} did not expand newline escape: got [$expanded]" ;;
esac
echo PASS
exit 0
