#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/case_compound_command_multiple_patterns
# In a case compound command, patterns separated by '|' match if any of the alternations match.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
match_type() {
    case "$1" in
        cat|dog|bird)
            printf 'animal\n'
            ;;
        car|truck)
            printf 'vehicle\n'
            ;;
        *)
            printf 'unknown\n'
            ;;
    esac
}
[ "$(match_type dog)" = "animal" ] || fail "match dog: want 'animal'"
[ "$(match_type truck)" = "vehicle" ] || fail "match truck: want 'vehicle'"
[ "$(match_type rock)" = "unknown" ] || fail "match rock: want 'unknown'"
echo PASS
exit 0
