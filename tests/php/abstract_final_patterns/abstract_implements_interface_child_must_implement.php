<?php
// vybe-test: php/abstract_final_patterns/abstract_implements_interface_child_must_implement
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

interface Runnable { public function run(): void; }
abstract class Task implements Runnable {
    public function schedule(): void { echo "scheduled:"; $this->run(); }
}
class PrintTask extends Task {
    public function run(): void { echo "running"; }
}
(new PrintTask())->schedule();

__vybe_check(ob_get_clean(), "scheduled:running");
