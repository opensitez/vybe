#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/coproc_keyword_starts_background_process
# The coproc reserved word launches a concurrent subshell and sets file descriptor array variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
coproc MYPROC { printf 'from_coproc\n'; }
read -r msg <&"${MYPROC[0]}"
wait "$MYPROC_PID"
[ "$msg" = "from_coproc" ] || fail "coproc message: want 'from_coproc', got [$msg]"
echo PASS
exit 0
