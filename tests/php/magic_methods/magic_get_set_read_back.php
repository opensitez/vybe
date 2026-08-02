<?php
// vybe-test: php/magic_methods/magic_get_set_read_back
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

class Registry {
    private array $store = [];
    public function __get($key) { return $this->store[$key] ?? null; }
    public function __set($key, $val) { $this->store[$key] = $val; }
    public function __isset($key) { return isset($this->store[$key]); }
}
$r = new Registry();
$r->a = 10;
$r->b = 20;
echo $r->a + $r->b;
echo isset($r->a) ? "set" : "not set";
echo isset($r->c) ? "set" : "not set";

__vybe_check(ob_get_clean(), "30setnot set");
