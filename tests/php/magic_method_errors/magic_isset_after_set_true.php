<?php
// vybe-test: php/magic_method_errors/magic_isset_after_set_true
// origin: languages/php/tests/php/test_magic_method_errors.rs

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
    private array $d = [];
    public function __set($k, $v) { $this->d[$k] = $v; }
    public function __isset($k) { return isset($this->d[$k]); }
}
$s = new Store();
$s->a = 1;
echo isset($s->a) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
