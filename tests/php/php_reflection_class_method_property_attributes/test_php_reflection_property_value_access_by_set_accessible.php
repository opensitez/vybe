<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_property_value_access_by_set_accessible
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs

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

class Account {
    private float $balance = 150.0;
}

$acc = new Account();
$rp = new ReflectionProperty(Account::class, "balance");
echo $rp->getValue($acc);

__vybe_check(ob_get_clean(), "150");
