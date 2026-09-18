#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/case_based_assignment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
classify() { case $1 in y|Y|yes) ans=yes ;; n|N|no) ans=no ;; *) ans=unknown ;; esac; }
classify Y; [ "$ans" = yes ] || fail "Y -> $ans"
classify no; [ "$ans" = no ] || fail "no -> $ans"
classify maybe; [ "$ans" = unknown ] || fail "maybe -> $ans"
echo PASS
exit 0
