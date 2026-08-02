<?php
// vybe-test: php/abstract_final_patterns/abstract_method_must_be_implemented_in_child
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

abstract class Formatter {
    abstract public function format(string $s): string;
    public function process(string $s): string { return $this->format($s); }
}
class UpperFormatter extends Formatter {
    public function format(string $s): string { return strtoupper($s); }
}
echo (new UpperFormatter())->process("hello"), "\n";

__vybe_check(ob_get_clean(), "HELLO");
