<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php84_property_hooks_get_set_syntax
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs

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

class Person {
    public string $first = "John";
    public string $last = "Doe";
    
    public string $fullName {
        get => "{$this->first} {$this->last}";
    }
}

$p = new Person();
echo $p->fullName;

__vybe_check(ob_get_clean(), "John Doe");
