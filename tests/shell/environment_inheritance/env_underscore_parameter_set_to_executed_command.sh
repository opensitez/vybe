#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_underscore_parameter_set_to_executed_command
# In an invoked child process, the '$_' environment variable contains the path of the executed binary.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
under=$( "$BASH" -c 'printf "%s\n" "$_"' )
case "$under" in
    *bash*) : ;;
    *) fail "\$_ in child does not match bash binary path: got [$under]" ;;
esac
echo PASS
exit 0
