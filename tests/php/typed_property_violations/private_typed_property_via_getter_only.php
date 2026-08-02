<?php
// vybe-test: php/typed_property_violations/private_typed_property_via_getter_only
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

class Wallet { private int $balance = 0; public function credit(int $n): void { $this->balance += $n; } public function total(): int { return $this->balance; } }
$w = new Wallet();
$w->credit(5);
echo $w->total();

__vybe_check(ob_get_clean(), "5");
