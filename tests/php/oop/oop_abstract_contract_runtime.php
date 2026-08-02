<?php
// vybe-test: php/oop/oop_abstract_contract_runtime
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

abstract class Writer {
    abstract public function write(string $text): string;
    public function prefix(): string { return 'x:' . $this->write('ok'); }
}
class UpperWriter extends Writer {
    public function write(string $text): string { return strtoupper($text); }
}
echo (new UpperWriter())->prefix();

__vybe_check(ob_get_clean(), "x:OK");
