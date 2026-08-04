<?php
// vybe-test: php/oop/enum_with_backed_value_and_match_runtime
// origin: languages/php/tests/php/test_oop.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

enum Role: string {
    case ADMIN = 'admin';
    case USER = 'user';
}
function label(Role $role): string {
    return match($role) {
        Role::ADMIN => 'A',
        Role::USER => 'U' };
}
echo label(Role::USER);

__vybe_check(ob_get_clean(), "U");
