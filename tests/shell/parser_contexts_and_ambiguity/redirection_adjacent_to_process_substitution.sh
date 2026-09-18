#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/redirection_adjacent_to_process_substitution
# A space between '<' (input redirection) and '<(...)' (process substitution) disambiguates the tokens.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r line < <(printf 'data_stream\n')
[ "$line" = "data_stream" ] || fail "redirection from process sub: want 'data_stream', got [$line]"
echo PASS
exit 0
