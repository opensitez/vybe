<?php
// vybe-test: php/magic_methods/magic_unset_removes_property
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

class Store {
    private array $map = ["k1" => "v1", "k2" => "v2", "k3" => "v3"];
    public function __isset($k) { return isset($this->map[$k]); }
    public function __unset($k) { unset($this->map[$k]); }
    public function count(): int { return count($this->map); }
}
$s = new Store();
echo $s->count();
unset($s->k2);
echo $s->count();
echo isset($s->k2) ? "yes" : "no";

__vybe_check(ob_get_clean(), "32no");
