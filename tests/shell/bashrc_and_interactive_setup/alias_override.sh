#!/usr/bin/env bash
# vybe-test: bash/bashrc_and_interactive_setup/alias_override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=9
/bin/bash -c '[[ $- != *i* ]]' || fail "non-interactive shell should not include -i"
/bin/bash -i -c '[[ $- == *i* ]]' || fail "interactive shell should include -i"
if (( IDX % 2 == 0 )); then
  rc_file="/tmp/vybe_bashrc_${IDX}_${BASHPID}.sh"
  printf '%s\n' 'BASHRC_MARK=ok' > "$rc_file"
  /bin/bash --rcfile "$rc_file" -i -c '[[ $BASHRC_MARK == ok ]]' || fail "--rcfile did not apply"
fi
echo PASS
exit 0
