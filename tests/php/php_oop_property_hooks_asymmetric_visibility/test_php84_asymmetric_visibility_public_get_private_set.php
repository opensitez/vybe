<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_asymmetric_visibility_public_get_private_set
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs

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

class BankAccount {
    public private(set) float $balance = 0.0;

    public function deposit(float $amount): void {
        $this->balance += $amount;
    }
}

$account = new BankAccount();
$account->deposit(250.0);
echo "Balance: {$account->balance}";

__vybe_check(ob_get_clean(), "Balance: 250");
