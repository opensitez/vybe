<?php
// vybe-test: php/oop_runtime/magic_set_stores_dynamic
// origin: languages/php/tests/php/test_oop_runtime.rs

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

class M {
    public array $d = [];
    public function __set(string $n, mixed $v): void { $this->d[$n] = $v; }
    public function read(): string { return $this->d['x']; }
}
$m = new M();
$m->x = 'ok';
echo $m->read();

__vybe_check(ob_get_clean(), "ok");
