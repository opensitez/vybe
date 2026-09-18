#!/usr/bin/env bash
# vybe-test: bash/truth_status_and_empty_values/status_zero_is_true_regardless_of_output
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (exit 0); then r=t; else r=f; fi
[ "$r" = t ] || fail "exit 0"
if (exit 3); then r=t; else r=f; fi
[ "$r" = f ] || fail "exit 3"
if :; then r=t; else r=f; fi
[ "$r" = t ] || fail ": is true"
if echo 0 >/dev/null; then r=t; else r=f; fi
[ "$r" = t ] || fail "printing 0 does not make a command false"
echo PASS
exit 0
