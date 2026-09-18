#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_multiline_value_inherited_by_child
# Environment variables containing literal line breaks are inherited intact without line truncation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export MULTILINE_DATA="line1"$'\n'"line2"$'\n'"line3"
child_count=$( "$BASH" -c 'count=0; while read -r _; do count=$((count+1)); done <<< "$MULTILINE_DATA"; printf "%s\n" "$count"' )
[ "$child_count" -eq 3 ] || fail "multiline environment variable line count in child: want 3, got $child_count"
echo PASS
exit 0
