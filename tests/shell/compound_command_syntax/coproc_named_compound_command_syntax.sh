#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/coproc_named_compound_command_syntax
# The coproc NAME { ...; } compound command executes a named coprocess and creates NAME array of fds.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
coproc WORKER {
    read -r req
    printf 'ack:%s\n' "$req"
}
printf 'ping\n' >&"${WORKER[1]}"
read -r res <&"${WORKER[0]}"
wait "$WORKER_PID"
[ "$res" = "ack:ping" ] || fail "coproc WORKER output: got [$res]"
echo PASS
exit 0
