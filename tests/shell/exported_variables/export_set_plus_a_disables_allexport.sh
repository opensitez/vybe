#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_set_plus_a_disables_allexport
# Disabling allexport via 'set +a' restores normal unexported variable assignment behavior.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -a
set +a
UNEXPORTED_POST="local_only_data"
child_val=$( "$BASH" -c 'printf "%s\n" "$UNEXPORTED_POST"' )
[ -z "$child_val" ] || fail "set +a failed to stop auto-exporting: child saw [$child_val]"
echo PASS
exit 0
