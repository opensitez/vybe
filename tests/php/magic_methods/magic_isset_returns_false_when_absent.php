<?php
// vybe-test: php/magic_methods/magic_isset_returns_false_when_absent
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Sparse {
    private array $vals = ["a" => 1, "c" => 3];
    public function __isset($k) { return array_key_exists($k, $this->vals); }
}
$s = new Sparse();
echo isset($s->a) ? "yes" : "no";
echo isset($s->b) ? "yes" : "no";
echo isset($s->c) ? "yes" : "no";

__vybe_check(ob_get_clean(), "yesnoyes");
