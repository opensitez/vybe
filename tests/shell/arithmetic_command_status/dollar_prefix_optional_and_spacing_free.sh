#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/dollar_prefix_optional_and_spacing_free
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=4
(( $x > 1 )) || fail "\$x form"
(( x > 1 )) || fail "bare name form"
((x>1)) || fail "no spaces form"
(( x>1&&x<9 )) || fail "compound without spaces"
echo PASS
exit 0
