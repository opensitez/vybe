<?php
// vybe-test: php/oop_interfaces/interface_inheritance_chain_runtime
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
interface Mid extends Base { public function mid(): string; }
interface Top extends Mid { public function top(): string; }
class Impl implements Top {
    public function base(): string { return 'b'; }
    public function mid(): string { return 'm'; }
    public function top(): string { return 't'; }
}
$obj = new Impl();
echo $obj->base() . $obj->mid() . $obj->top();

__vybe_check(ob_get_clean(), "bmt");
