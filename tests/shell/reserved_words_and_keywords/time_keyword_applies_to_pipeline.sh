#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/time_keyword_applies_to_pipeline
# The time reserved word times an entire pipeline and passes through the pipeline's exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
time { printf 'pipeline_out\n' | cat; } >/dev/null 2>/dev/null
st=$?
[ "$st" -eq 0 ] || fail "time pipeline status: want 0, got $st"
echo PASS
exit 0
