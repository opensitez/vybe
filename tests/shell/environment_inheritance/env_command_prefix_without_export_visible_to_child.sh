#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_command_prefix_without_export_visible_to_child
# A temporary prefix assignment is placed into the child process environment even without explicit 'export'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset AD_HOC_KEY
child_val=$( AD_HOC_KEY="passed_in_prefix" "$BASH" -c 'printf "%s\n" "$AD_HOC_KEY"' )
[ "$child_val" = "passed_in_prefix" ] || fail "prefix variable not seen in child: got [$child_val]"
[ -z "$AD_HOC_KEY" ] || fail "prefix variable leaked to parent shell"
echo PASS
exit 0
