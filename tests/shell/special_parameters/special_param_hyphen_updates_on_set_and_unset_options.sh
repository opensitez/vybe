#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_hyphen_updates_on_set_and_unset_options
# Modifying shell options via 'set -f' or 'set +f' dynamically updates $- string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -f
case "$-" in
    *f*) : ;;
    *) fail "\$- missing flag 'f' after 'set -f': got [$-]" ;;
esac
set +f
case "$-" in
    *f*) fail "\$- still contains flag 'f' after 'set +f': got [$-]" ;;
esac
echo PASS
exit 0
