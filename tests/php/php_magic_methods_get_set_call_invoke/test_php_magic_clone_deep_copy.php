<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_clone_deep_copy
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs

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

class Address {
    public string $city = "NY";
}

class User {
    public Address $address;
    public function __construct() {
        $this->address = new Address();
    }
    public function __clone() {
        $this->address = clone $this->address;
    }
}

$u1 = new User();
$u2 = clone $u1;
$u2->address->city = "LA";
echo $u1->address->city . " vs " . $u2->address->city;

__vybe_check(ob_get_clean(), "NY vs LA");
