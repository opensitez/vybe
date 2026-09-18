#!/usr/bin/env bash
# vybe-test: bash/process_substitution/output_form_feeds_a_command
# >(list) expands to a path; whatever is written there becomes list's stdin.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$( echo data > >(read -r l; echo "got:$l") )
[ "$out" = got:data ] || fail "want [got:data] got [$out]"
echo PASS
exit 0
