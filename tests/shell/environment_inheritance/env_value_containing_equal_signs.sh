#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_value_containing_equal_signs
# Environment variable values containing '=' characters are inherited faithfully across processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export CONNECTION_STR="key1=val1;key2=val2=extra"
child_val=$( "$BASH" -c 'printf "%s\n" "$CONNECTION_STR"' )
[ "$child_val" = "key1=val1;key2=val2=extra" ] || fail "equal signs in value corrupted: got [$child_val]"
echo PASS
exit 0
