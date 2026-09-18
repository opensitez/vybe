#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_variable_preserving_newlines
# Exported variables pass embedded literal newline characters intact across process execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export MULTILINE_EXP="row1
row2
row3"
line_count=$( "$BASH" -c 'count=0; while read -r l; do count=$((count+1)); done <<< "$MULTILINE_EXP"; echo "$count"' )
[ "$line_count" -eq 3 ] || fail "multiline exported variable lost lines in child: got $line_count"
echo PASS
exit 0
