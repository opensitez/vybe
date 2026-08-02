<?php
// vybe-test: php/clone_patterns/clone_magic_method_deep_clones_nested_object
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Address { public string $city; public function __construct(string $city) { $this->city = $city; } }
class Person {
    public Address $address;
    public function __construct(public string $name, Address $addr) { $this->address = $addr; }
    public function __clone() { $this->address = clone $this->address; }
}
$alice = new Person("Alice", new Address("London"));
$bob = clone $alice;
$bob->address->city = "Paris";
echo $alice->address->city . ',' . $bob->address->city;

__vybe_check(ob_get_clean(), "London,Paris");
