#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_declare_dash_p_shows_reference
# The 'declare -p ref' command displays the nameref attribute and the name of its target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source_item="val"
declare -n ref=source_item
desc=$(declare -p ref)
case "$desc" in
    *'declare -n ref="source_item"'*) : ;;
    *) fail "declare -p output did not show nameref assignment: got [$desc]" ;;
esac
echo PASS
exit 0
