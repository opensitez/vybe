<?php
// vybe-test: php/oop_runtime/magic_unset_clears
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
    public array $d = ['a' => 1];
    public function __unset(string $k): void { unset($this->d[$k]); }
}
$m = new M();
unset($m->a);
echo isset($m->d['a']) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "no");
