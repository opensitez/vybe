#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_dash_p_lists_exported_variables
# The 'export -p' builtin lists all exported variables in reusable declaration syntax.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export PROBE_EXPORT_P="probe_val_99"
listing=$(export -p)
case "$listing" in
    *"PROBE_EXPORT_P="*) : ;;
    *) fail "export -p missing declared export variable" ;;
esac
echo PASS
exit 0
