#!/usr/bin/env bash
# vybe-test: bash/dotglob_behavior/dot_and_dotdot_are_skipped_by_globskipdots
# globskipdots (on by default since bash 5.2) keeps . and .. out of .*;
# turning it off brings them back.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > .hidden
shopt -s globskipdots
[ "$(echo .*)" = .hidden ] || fail "with globskipdots: got [$(echo .*)]"
shopt -u globskipdots
[ "$(echo .*)" = ". .. .hidden" ] || fail "without globskipdots: got [$(echo .*)]"
echo PASS
exit 0
