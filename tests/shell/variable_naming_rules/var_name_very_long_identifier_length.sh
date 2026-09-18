#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_very_long_identifier_length
# Bash supports arbitrarily long variable identifier names (e.g. 100+ alphanumeric characters).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
long_identifier_name_with_many_descriptive_parts_to_verify_that_bash_handles_arbitrarily_long_names_without_truncating_or_crashing_the_parser_state="payload"
[ "$long_identifier_name_with_many_descriptive_parts_to_verify_that_bash_handles_arbitrarily_long_names_without_truncating_or_crashing_the_parser_state" = "payload" ] || fail "long identifier failed"
echo PASS
exit 0
