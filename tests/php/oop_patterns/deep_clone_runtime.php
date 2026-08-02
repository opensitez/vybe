<?php
// vybe-test: php/oop_patterns/deep_clone_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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
    public function __construct(public string $city) {}
}
class Person {
    public function __construct(public string $name, public Address $address) {}
    public function __clone(): void {
        $this->address = clone $this->address;
    }
}
$original = new Person('Alice', new Address('Paris'));
$copy = clone $original;
$copy->address->city = 'London';
echo $original->address->city . '|' . $copy->address->city;

__vybe_check(ob_get_clean(), "Paris|London");
