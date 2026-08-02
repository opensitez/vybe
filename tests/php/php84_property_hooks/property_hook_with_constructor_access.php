<?php
// vybe-test: php/php84_property_hooks/property_hook_with_constructor_access
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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

class Email {
    public string $address {
        set(string $v) {
            if (!str_contains($v, '@')) throw new \InvalidArgumentException("Invalid email");
            $this->address = strtolower($v);
        }
    }
    public function __construct(string $addr) { $this->address = $addr; }
}
$e = new Email("Alice@Example.COM");
echo $e->address;

__vybe_check(ob_get_clean(), "alice@example.com");
