#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_terminates_hash_comment
# A newline character terminates a comment introduced by '#' and permits code execution on the next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
# This is a comment line ending with newline
actual="executed"
[ "$actual" = "executed" ] || fail "comment termination failed: got [$actual]"
echo PASS
exit 0
