#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_parameter_default_expansion
# In unquoted parameter default expansions ${var:-word}, tilde expansion is performed on the default word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
unset unassigned_folder
# Unquoted default word undergoes tilde expansion
resolved=${unassigned_folder:-~/backup_dir}
[ "$resolved" = "$HOME/backup_dir" ] || fail "tilde in parameter default failed: want [$HOME/backup_dir], got [$resolved]"

# Double-quoted parameter expansion suppresses tilde expansion
quoted_resolved="${unassigned_folder:-~/backup_dir}"
[ "$quoted_resolved" = "~/backup_dir" ] || fail "double quotes should suppress tilde expansion: got [$quoted_resolved]"
echo PASS
exit 0
