#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_set_dash_a_allexport_automatically_exports
# With 'set -a' (allexport) enabled, all subsequently assigned variables are automatically exported.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -a
AUTOMATIC_EXPORT="exported_without_explicit_keyword"
set +a
child_val=$( "$BASH" -c 'printf "%s\n" "$AUTOMATIC_EXPORT"' )
[ "$child_val" = "exported_without_explicit_keyword" ] || fail "allexport failed: got [$child_val]"
echo PASS
exit 0
