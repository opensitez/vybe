<?php
// vybe-test: php/inheritance_patterns/polymorphic_dispatch
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

abstract class Payment { abstract public function pay(float $amount): string; }
class CreditCard extends Payment { public function pay(float $a): string { return "CC:$a"; } }
class PayPal extends Payment { public function pay(float $a): string { return "PP:$a"; } }
$payments = [new CreditCard, new PayPal, new CreditCard];
echo implode(',', array_map(fn($p) => $p->pay(10.0), $payments)), "\n";

__vybe_check(ob_get_clean(), "CC:10,PP:10,CC:10");
