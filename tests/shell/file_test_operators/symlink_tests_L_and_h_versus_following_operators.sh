#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/symlink_tests_L_and_h_versus_following_operators
# -L and -h look at the link itself; every other operator follows it, so a
# dangling link is -L but not -e.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > target; ln -s target link; ln -s nowhere dangling
[ -L link ] && [ -h link ] || fail "-L/-h on link"
[ -L target ] && fail "-L on a regular file"
[ -f link ] || fail "-f follows the link"
[ -L dangling ] || fail "-L on dangling link"
[ -e dangling ] && fail "-e follows the dangling link and must be false"
echo PASS
exit 0
