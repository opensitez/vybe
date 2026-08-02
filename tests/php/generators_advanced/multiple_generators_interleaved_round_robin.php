<?php
// vybe-test: php/generators_advanced/multiple_generators_interleaved_round_robin
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function taskA() { yield "A1"; yield "A2"; yield "A3"; }
function taskB() { yield "B1"; yield "B2"; }
$gens = [taskA(), taskB()];
$output = [];
$alive = true;
while ($alive) {
    $alive = false;
    foreach ($gens as $g) {
        if ($g->valid()) {
            $output[] = $g->current();
            $g->next();
            $alive = true;
        }
    }
}
echo implode(",", $output);

__vybe_check(ob_get_clean(), "A1,B1,A2,B2,A3");
