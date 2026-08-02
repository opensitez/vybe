<?php
// vybe-test: php/php_oop_property_hooks_get_set_syntax/test_property_hooks_set_transformation
// origin: languages/php/tests/php/test_php_oop_property_hooks_get_set_syntax.rs

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

class SanitizedUser {
    private string $_email = '';

    public string $email {
        get => $this->_email;
        set => $this->_email = strtolower(trim($value));
    }
}
$u = new SanitizedUser();
$u->email = '   USER@EXAMPLE.COM   ';
echo $u->email, "\n";

__vybe_check(ob_get_clean(), "user@example.com");
