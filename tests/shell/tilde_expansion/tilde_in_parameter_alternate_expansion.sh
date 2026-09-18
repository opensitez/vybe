#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_parameter_alternate_expansion
# The word of ${var:+word} undergoes tilde expansion, but only when the whole
# expansion is unquoted: inside double quotes no tilde expansion happens, so
# "${var:+~/x}" keeps a literal ~.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
flag="enabled"
unquoted=${flag:+~/alt_dir}
[ "$unquoted" = "$HOME/alt_dir" ] || fail "unquoted alternate: want [$HOME/alt_dir] got [$unquoted]"
quoted="${flag:+~/alt_dir}"
[ "$quoted" = "~/alt_dir" ] || fail "double-quoted alternate must keep literal ~, got [$quoted]"
echo PASS
exit 0
