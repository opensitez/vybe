#!/usr/bin/env bash
# vybe-test: bash/nocaseglob_behavior/result_keeps_actual_filename_case
# A matched name is returned as stored on disk. A word without any glob
# character is not a pattern and is left exactly as typed, even under nocaseglob.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > MiXeD.dat
shopt -s nocaseglob
for f in mixed.d?t; do got=$f; done
[ "$got" = MiXeD.dat ] || fail "pattern match: got [$got]"
for f in mixed.dat; do plain=$f; done
[ "$plain" = mixed.dat ] || fail "non-pattern word must stay literal: got [$plain]"
echo PASS
exit 0
