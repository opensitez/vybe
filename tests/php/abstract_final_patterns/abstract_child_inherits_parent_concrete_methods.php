<?php
// vybe-test: php/abstract_final_patterns/abstract_child_inherits_parent_concrete_methods
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

abstract class Writer {
    public function writeLine(string $s): void { echo $s . "\n", "\n"; }
    abstract public function target(): string;
}
abstract class NetworkWriter extends Writer {
    public function prefix(): string { return "[net] "; }
}
class HttpWriter extends NetworkWriter {
    public function target(): string { return "HTTP"; }
    public function write(string $msg): void { $this->writeLine($this->prefix() . $msg); }
}
(new HttpWriter())->write("hello");

__vybe_check(ob_get_clean(), "[net] hello");
