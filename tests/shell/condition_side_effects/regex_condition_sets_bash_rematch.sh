#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/regex_condition_sets_bash_rematch
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if [[ key=value =~ ^([^=]+)=(.*)$ ]]; then k=${BASH_REMATCH[1]}; v=${BASH_REMATCH[2]}; fi
[ "$k" = key ] && [ "$v" = value ] || fail "k=[$k] v=[$v]"
echo PASS
exit 0
