#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_pointing_to_associative_array_key
# A nameref variable can point directly to a specific key of an associative array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [theme]="dark" )
declare -n current_theme='map[theme]'
[ "$current_theme" = "dark" ] || fail "reading associative array key through nameref failed"
current_theme="light"
[ "${map[theme]}" = "light" ] || fail "associative array key not mutated through nameref: got [${map[theme]}]"
echo PASS
exit 0
