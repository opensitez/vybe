<?php
// vybe-test: php/modern_php_deep/nullsafe_deep_chain
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

class Country {
    public function __construct(public string $name) {}
}
class Address {
    public ?Country $country;
    public function __construct(?Country $country) {
        $this->country = $country;
    }
}
class User {
    public ?Address $address;
    public function __construct(?Address $address) {
        $this->address = $address;
    }
}
$u1 = new User(new Address(new Country("USA")));
$u2 = new User(new Address(null));
$u3 = new User(null);
echo $u1?->address?->country?->name ?? "unknown";
echo $u2?->address?->country?->name ?? "unknown";
echo $u3?->address?->country?->name ?? "unknown";

__vybe_check(ob_get_clean(), "USAunknownunknown");
