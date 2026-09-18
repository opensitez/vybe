#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/time_keyword_reports_on_stderr
# time is a reserved word: it times a whole pipeline and prints TIMEFORMAT to
# stderr; the pipeline's stdout is untouched.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
TIMEFORMAT=timing-done
out=$( { time echo payload; } 2>/dev/null )
[ "$out" = payload ] || fail "stdout: want [payload] got [$out]"
err=$( { time echo payload; } 2>&1 >/dev/null )
[ "$err" = timing-done ] || fail "stderr: want [timing-done] got [$err]"
echo PASS
exit 0
