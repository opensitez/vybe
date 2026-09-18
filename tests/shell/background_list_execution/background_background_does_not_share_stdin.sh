#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_does_not_share_stdin
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r x <<'EOF_IN'
literal
EOF_IN
( read -r bg </dev/null ) &
pid=$!
wait "$pid"
[ "$x" = "literal" ] || fail "background read should not consume parent stdin"
echo PASS
exit 0
