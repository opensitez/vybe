#!/usr/bin/env bash
# vybe-test: bash/dotglob_behavior/dotglob_applies_within_subdirectory_component
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
mkdir sub; : > sub/a; : > sub/.h
[ "$(count sub/*)" = 1 ] || fail "default: want 1 got $(count sub/*)"
shopt -s dotglob
[ "$(count sub/*)" = 2 ] || fail "dotglob: want 2 got $(count sub/*)"
echo PASS
exit 0
