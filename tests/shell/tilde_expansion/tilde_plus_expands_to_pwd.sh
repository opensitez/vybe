#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_plus_expands_to_pwd
# An unquoted '~+' expands to the value of the PWD variable (current working directory).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
PWD="/custom/working/directory"
res=~+
[ "$res" = "/custom/working/directory" ] || fail "~+ expansion failed: want '/custom/working/directory', got [$res]"
echo PASS
exit 0
