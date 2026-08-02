<?php
// vybe-test: php/oop_interfaces/instanceof_checks_interface
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Tagged { public function tag(): string; }
class Item implements Tagged { public function tag(): string { return 'item'; } }
$o = new Item;
echo ($o instanceof Tagged) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
