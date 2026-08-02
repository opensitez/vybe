<?php
// vybe-test: php/intersection_types/intersection_type_narrower_is_subtype_of_single_interface
// origin: languages/php/tests/php/test_intersection_types.rs

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

interface Worker { public function work(): string; }
interface Reporter { public function report(): string; }
class FullEmployee implements Worker, Reporter {
    public function work(): string { return "working"; }
    public function report(): string { return "reporting"; }
}
function getWorker(): Worker { return new FullEmployee(); }
$w = getWorker();
echo $w->work();

__vybe_check(ob_get_clean(), "working");
