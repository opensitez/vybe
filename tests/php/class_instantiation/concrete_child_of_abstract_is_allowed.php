<?php
// vybe-test: php/class_instantiation/concrete_child_of_abstract_is_allowed
// origin: languages/php/tests/php/test_class_instantiation.rs

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

abstract class Worker { abstract public function run(): string; }
class Job extends Worker { public function run(): string { return 'done'; } }
echo (new Job())->run();

__vybe_check(ob_get_clean(), "done");
