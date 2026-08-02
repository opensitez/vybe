<?php
// vybe-test: php/modern_php_deep/nullsafe_chain_succeeding
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

class Address {
    public function __construct(public string $city) {}
    public function getCity(): string { return $this->city; }
}
class User {
    public ?Address $address;
    public function __construct(?Address $addr) { $this->address = $addr; }
}
$u = new User(new Address("Paris"));
echo $u?->address?->getCity() ?? "unknown";

__vybe_check(ob_get_clean(), "Paris");
