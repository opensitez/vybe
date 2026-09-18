#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_associative_array_declaration_dash_A
# Declaring 'local -A map' creates an associative array scoped exclusively to the function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
map_fn() {
    local -A config=( [host]="localhost" [port]="8080" )
    [ "${config[host]}" = "localhost" ] || fail "config[host]: want 'localhost', got [${config[host]}]"
    [ "${config[port]}" = "8080" ] || fail "config[port]: want '8080', got [${config[port]}]"
}
map_fn
[[ ! -v config ]] || fail "local associative array config should not exist in outer scope"
echo PASS
exit 0
