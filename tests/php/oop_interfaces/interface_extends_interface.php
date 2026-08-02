<?php
// vybe-test: php/oop_interfaces/interface_extends_interface
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

interface Base { public function base(): string; }
interface Extended extends Base { public function extra(): string; }
class Impl implements Extended {
    public function base(): string { return 'base'; }
    public function extra(): string { return 'extra'; }
}
$o = new Impl;
echo $o->base() . ',' . $o->extra();

__vybe_check(ob_get_clean(), "base,extra");
