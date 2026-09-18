#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/indirect_parameter_name_resolution
# The ${!name} expansion dynamically resolves the variable identifier held inside another variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
actual_data="payload"
pointer="actual_data"
resolved="${!pointer}"
[ "$resolved" = "payload" ] || fail "indirect resolution: want 'payload', got [$resolved]"
echo PASS
exit 0
