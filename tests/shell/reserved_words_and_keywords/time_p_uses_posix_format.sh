#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/time_p_uses_posix_format
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
TIMEFORMAT=ignored-when-p
err=$( { time -p true; } 2>&1 )
[[ $err == real\ *user\ *sys\ * ]] || fail "want real/user/sys lines got [$err]"
[[ $err != *ignored-when-p* ]] || fail "-p must override TIMEFORMAT"
echo PASS
exit 0
