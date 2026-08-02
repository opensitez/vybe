<?php
// vybe-test: php/modern_php_deep/enum_used_as_array_key_via_value
// origin: languages/php/tests/php/test_modern_php_deep.rs

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
    case Admin = "admin";
    case User  = "user";
    case Guest = "guest";
}
$permissions = [
    Role::Admin->value => ["read", "write", "delete"],
    Role::User->value  => ["read", "write"],
    Role::Guest->value => ["read"],
];
$role = Role::User;
echo count($permissions[$role->value]);
echo $permissions[Role::Guest->value][0];

__vybe_check(ob_get_clean(), "2read");
