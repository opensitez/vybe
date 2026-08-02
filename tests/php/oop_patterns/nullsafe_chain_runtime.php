<?php
// vybe-test: php/oop_patterns/nullsafe_chain_runtime
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

class Street { public function __construct(public string $name) {} }
class Address { public function __construct(public ?Street $street = null) {} }
class User {
    public function __construct(public ?Address $address = null) {}
}
echo (new User(new Address(new Street('Main')))?->address?->street?->name ?? 'none'), '|', (new User())->address?->street?->name ?? 'none';

__vybe_check(ob_get_clean(), "Main|none");
