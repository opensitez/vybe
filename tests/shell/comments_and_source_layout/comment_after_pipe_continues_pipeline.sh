#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_after_pipe_continues_pipeline
# A comment on the same line as a trailing pipe does not break pipeline continuation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=$(printf '%s' "routed" | # send output to next stage
    cat)
[ "$res" = "routed" ] || fail "comment after pipe: want 'routed', got [$res]"
echo PASS
exit 0
