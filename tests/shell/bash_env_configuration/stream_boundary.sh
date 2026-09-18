#!/usr/bin/env bash
# vybe-test: bash/bash_env_configuration/stream_boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=12
env_file="/tmp/vybe_env_${IDX}_${BASHPID}.sh"
printf '%s\n' 'BASH_ENV_MARK=from_bash_env' > "$env_file"
BASH_ENV="$env_file" bash -c '[[ $BASH_ENV_MARK = from_bash_env ]]' || fail "BASH_ENV file not sourced"
[[ -z ${BASH_ENV_MARK+x} ]] || fail "BASH_ENV marker leaked to parent shell"
if (( IDX % 2 == 0 )); then
  BASH_ENV="$env_file" bash -c '[[ $BASH_ENV_MARK == from_bash_env ]] || exit 1'
fi
echo PASS
exit 0
