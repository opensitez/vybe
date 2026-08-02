<?php
// vybe-test: php/oop/multiple_inheritance_interfaces_runtime
// origin: languages/php/tests/php/test_oop.rs

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

interface Logger { public function log(): string; }
interface Formatter { public function format(string $s): string; }
class Item implements Logger, Formatter {
    public function log(): string { return 'log'; }
    public function format(string $s): string { return strrev($s); }
}
$i = new Item();
echo $i->log(), '|', $i->format('ab');

__vybe_check(ob_get_clean(), "log|ba");
