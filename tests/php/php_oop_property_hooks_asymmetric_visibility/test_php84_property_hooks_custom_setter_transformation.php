<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_property_hooks_custom_setter_transformation
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs

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

class UserProfile {
    public string $email {
        set => strtolower(trim($value));
    }
}

$u = new UserProfile();
$u->email = "  ALICE@Example.COM  ";
echo $u->email;

__vybe_check(ob_get_clean(), "alice@example.com");
