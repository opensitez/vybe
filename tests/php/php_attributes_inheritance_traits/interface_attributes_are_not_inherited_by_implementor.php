<?php
// vybe-test: php/php_attributes_inheritance_traits/interface_attributes_are_not_inherited_by_implementor
// origin: languages/php/tests/php/test_php_attributes_inheritance_traits.rs

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

#[Attribute]
class Contract {}
#[Contract]
interface Payable {}
class Invoice implements Payable {}
echo count((new ReflectionClass(Invoice::class))->getAttributes(Contract::class));

__vybe_check(ob_get_clean(), "0");
