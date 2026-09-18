#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/literal_hash_and_percent_in_pattern
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='#a%'
[ "${x#\#}" = 'a%' ] || fail "escaped #: got [${x#\#}]"
[ "${x%\%}" = '#a' ] || fail "escaped %: got [${x%\%}]"
[ "${x#'#'}" = 'a%' ] || fail "quoted #: got [${x#'#'}]"
echo PASS
exit 0
