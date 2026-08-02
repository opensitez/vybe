<?php
// vybe-test: php/magic_methods/magic_isset_returns_true_for_existing
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

class Box {
    private array $data = ["x" => 1];
    public function __isset($name) {
        return isset($this->data[$name]);
    }
    public function __unset($name) {
        unset($this->data[$name]);
    }
}
$b = new Box();
echo isset($b->x) ? "yes" : "no";
echo isset($b->y) ? "yes" : "no";
unset($b->x);
echo isset($b->x) ? "yes" : "no";

__vybe_check(ob_get_clean(), "yesnono");
