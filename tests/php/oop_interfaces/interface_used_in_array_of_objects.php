<?php
// vybe-test: php/oop_interfaces/interface_used_in_array_of_objects
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Worker { public function work(): int; }
class DevWorker implements Worker { public function work(): int { return 8; } }
class OpWorker implements Worker { public function work(): int { return 10; } }
$workers = [new DevWorker, new OpWorker, new DevWorker];
echo array_sum(array_map(fn(Worker $w) => $w->work(), $workers));

__vybe_check(ob_get_clean(), "26");
