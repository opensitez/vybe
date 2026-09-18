#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_colon_separated_path_assignment
# Path-like variables with colon delimiters preserve tilde expansion after colons without quotes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
mypath=~/bin:~/opt
[ "$mypath" = "$HOME/bin:$HOME/opt" ] || fail "tilde expansion after colon in assignment failed: got [$mypath]"
echo PASS
exit 0
