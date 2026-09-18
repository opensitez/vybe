#!/usr/bin/env bash
# vybe-test: bash/nocaseglob_behavior/does_not_affect_case_or_double_bracket_matching
# nocaseglob is for filenames only; string pattern matching needs nocasematch.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s nocaseglob
case A in a) r=ci;; *) r=cs;; esac
[ "$r" = cs ] || fail "case must stay case sensitive"
[[ A == a ]] && fail "[[ must stay case sensitive"
echo PASS
exit 0
