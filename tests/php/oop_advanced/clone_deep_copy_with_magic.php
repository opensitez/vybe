<?php
// vybe-test: php/oop_advanced/clone_deep_copy_with_magic
// origin: languages/php/tests/php/test_oop_advanced.rs

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
    public Address $address;
    public function __construct(public string $name, string $city) {
        $this->address = new Address($city);
    }
    public function __clone() {
        $this->address = clone $this->address;
    }
}
$alice = new Person("Alice", "Paris");
$bob = clone $alice;
$bob->name = "Bob";
$bob->address->city = "London";
echo $alice->name . ":" . $alice->address->city, "\n";
echo $bob->name . ":" . $bob->address->city, "\n";

__vybe_check(ob_get_clean(), "Alice:Paris\nBob:London");
